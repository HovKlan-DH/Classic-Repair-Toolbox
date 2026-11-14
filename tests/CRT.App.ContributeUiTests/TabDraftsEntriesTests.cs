using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE LEFT COLUMN'S ENTRIES AND "UI" OR "TABLE" (TabDrafts.Entries.cs - owner requests, 2026-11-03: the
// component's pages under "Components" in place of tabs, and "when I select e.g. 'Board: Schematic images'
// then I will see the table for 'Board schematics' in the content area").
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsEntriesTests : IDisposable
{
    private const string BoardA = "Commodore/C64/250407/board.json";
    private const string BoardB = "Commodore/C64/250425/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsEntriesTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry Entry(string board = "250407", string excelDataFile = BoardA) => new()
    {
        HardwareName = "Commodore 64",
        BoardName = board,
        ExcelDataFile = excelDataFile,
    };

    private static BoardData Board() => new()
    {
        Regions = ["PAL", "NTSC"],
        Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" }],
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU", Package = "DIP-40" },
            new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "VIC", Region = "PAL", TechnicalNameOrValue = "6569" },
            new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "VIC", Region = "NTSC", TechnicalNameOrValue = "6567" },
            new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", FriendlyName = "Filter" },
        ],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/a.png" },
            new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "Scope baseline/b.png" },
        ],
        ComponentPins = [new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" }],
        BoardLinks = [new BoardLinkEntry { Category = "Web", Name = "Forum", Url = "https://example.com" }],
    };

    private static void WriteDraft(string excelDataFile = BoardA, BoardData? board = null)
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board ?? Board());

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker { BoardKey = excelDataFile });
    }

    // The tab on a board, with the list on the far left that Main has for it.
    private static (TabDrafts Tab, ListBox List) Tab()
    {
        HardwareBoardEntry entry = Entry();

        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [entry, Entry("250425", BoardB)],
            PublishedBoardOverrideForTests = _ => Board(),
        };

        tab.RefreshDrafts();
        tab.ShowBoard(entry, Board(), "PAL", null);

        var list = new ListBox { SelectionMode = SelectionMode.Single, ItemsSource = tab.ComponentItems };
        tab.UseComponentList(list);

        return (tab, list);
    }

    private static string Count(TabDrafts tab, string name) => tab.GetControl<TextBlock>(name).Text ?? string.Empty;

    private static void Click(TabDrafts tab, string name) =>
        tab.GetControl<Button>(name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    private static async Task Until(Func<bool> condition, string what)
    {
        for (int attempt = 0; attempt < 300 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(condition(), what);
    }

    // ###########################################################################################
    // Owner report, 2026-11-11: "if I have too many or too long names for the regions, then it will take all space in the
    // menu, and I cannot see the "Regions" entry to click" (the Open128-2 board). Six long region names: "Regions" stays whole,
    // and the names are cut short in the room left beside it, never past the entry's edge.
    // ###########################################################################################
    [Fact]
    public void Long_region_names_never_squeeze_the_Regions_entry_out_of_the_column()
    {
        UiTest.Run(() =>
        {
            BoardData board = Board().WithRegions(["Europe PAL-B 50Hz", "North America NTSC", "Japan NTSC-J", "Brazil PAL-M", "Argentina PAL-N", "Australia"]);

            HardwareBoardEntry entry = Entry();
            var tab = new TabDrafts
            {
                HardwareBoardsOverrideForTests = [entry],
                PublishedBoardOverrideForTests = _ => board,
            };

            tab.RefreshDrafts();
            tab.ShowBoard(entry, board, "Japan NTSC-J", null);

            var window = new Window { Content = tab, Width = 1000, Height = 800 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Button regions = tab.GetControl<Button>("RegionsNavButton");
            TextBlock name = regions.GetVisualDescendants().OfType<TextBlock>().First(text => text.Text == "Regions");
            TextBlock names = tab.GetControl<TextBlock>("RegionsCountText");

            Assert.StartsWith("Europe PAL-B 50Hz, North America NTSC", names.Text, StringComparison.Ordinal);

            name.Measure(Size.Infinity);
            Assert.True(name.Bounds.Width >= name.DesiredSize.Width - 0.5, $"\"Regions\" is [{name.Bounds.Width}] wide, needs [{name.DesiredSize.Width}]");
            Assert.Equal(TextTrimming.CharacterEllipsis, names.TextTrimming);

            Point namesEnd = names.TranslatePoint(new Point(names.Bounds.Width, 0), regions)!.Value;
            Assert.True(namesEnd.X <= regions.Bounds.Width + 0.5, $"the names end at [{namesEnd.X}] of [{regions.Bounds.Width}]");

            window.Close();
        });
    }

    // ###########################################################################################
    // A component "Board: Regions" names (owner request, 2026-11-11: "so I can find these and change them") opens on
    // "Basics", picked in the list, from whatever was open - the region switched to the one it was named under.
    // ###########################################################################################
    [Fact]
    public async Task A_component_a_region_is_applied_on_opens_on_Basics_in_that_region()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            await tab.OpenBoardPartAsync(PaneBoardPart.Regions);
            await tab.OpenComponentInRegionAsync("U1", "NTSC");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            Assert.Equal("U1", tab.ComponentPaneForTests.Label);
            Assert.Equal(ComponentPage.Basics, tab.ComponentPaneForTests.Page);
            Assert.Equal("BasicsNavButton", tab.SelectedNavEntryForTests);
            Assert.Equal("6567", tab.ComponentPaneForTests.GetControl<TextBox>("TechnicalNameBox").Text);
            Assert.Equal("U1", (list.SelectedItem as ComponentListItem)?.BoardLabel);
        });
    }

    // ---- a component's pages -------------------------------------------------------------------------

    // Owner request, 2026-11-03: "If a component is not selected, then in its content area then state that
    // the user can use and select a component from the normal left-side component list".
    [Fact]
    public async Task With_no_component_picked_a_page_says_where_to_pick_one_and_a_pick_opens_on_it()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            await tab.ShowEntryAsync(ContributeEntry.ComponentPinout);

            Assert.True(tab.IsNoComponentPageShownForTests);
            Assert.False(tab.IsComponentOpen);
            Assert.False(tab.GetControl<StackPanel>("StartPanel").IsVisible);
            Assert.Equal("PinoutNavButton", tab.SelectedNavEntryForTests);
            Assert.Contains("\"Pinout\"", tab.NoComponentTextForTests, StringComparison.Ordinal);
            Assert.Contains("component list", tab.NoComponentTextForTests, StringComparison.Ordinal);

            // Nothing picked, nothing counted.
            Assert.Equal(string.Empty, Count(tab, "PinoutCountText"));

            tab.PickComponentForTests("U8");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            Assert.False(tab.IsNoComponentPageShownForTests);
            Assert.Equal(ComponentPage.Pinout, tab.ComponentPaneForTests.Page);
            Assert.Equal("PinoutNavButton", tab.SelectedNavEntryForTests);
            Assert.Equal("U8", (list.SelectedItem as ComponentListItem)?.BoardLabel);
        });
    }

    // ###########################################################################################
    // "Edit component" on a popup opens the component on "Basics", whichever page the pane was on (owner request,
    // 2026-11-07: "it should go to the "Basics" menu for that component") - its entry shown selected.
    // ###########################################################################################
    [Fact]
    public async Task Edit_component_opens_the_component_on_Basics_whichever_page_was_shown()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowEntryAsync(ContributeEntry.ComponentPinout);
            tab.PickComponentForTests("U8");
            await Until(() => tab.IsComponentOpen, "The component never opened.");
            Assert.Equal(ComponentPage.Pinout, tab.ComponentPaneForTests.Page);

            await tab.EditComponentAsync("U8");

            Assert.True(tab.IsComponentOpen);
            Assert.Equal(ComponentPage.Basics, tab.ComponentPaneForTests.Page);
            Assert.Equal("BasicsNavButton", tab.SelectedNavEntryForTests);
        });
    }

    // Overview closes it again, as anything else open.
    [Fact]
    public async Task Overview_goes_back_from_the_page_with_no_component()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowEntryAsync(ContributeEntry.ComponentLinks);
            await tab.ShowEntryAsync(ContributeEntry.Overview);

            Assert.False(tab.IsNoComponentPageShownForTests);
            Assert.True(tab.GetControl<StackPanel>("StartPanel").IsVisible);
            Assert.Equal("StatusNavButton", tab.SelectedNavEntryForTests);
        });
    }

    [Fact]
    public async Task Each_entry_shows_the_page_of_the_component_picked_and_counts_what_it_holds()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            tab.PickComponentForTests("U8");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            Assert.Equal("BasicsNavButton", tab.SelectedNavEntryForTests);
            Assert.Equal("1", Count(tab, "BasicsCountText"));
            Assert.Equal("1", Count(tab, "PinoutCountText"));
            Assert.Equal("1", Count(tab, "ComponentPicturesCountText"));
            Assert.Equal("1", Count(tab, "BaselinesCountText"));
            Assert.Equal("0", Count(tab, "ReplacementsCountText"));

            await tab.ShowEntryAsync(ContributeEntry.ComponentBaselines);

            Assert.Equal(ComponentPage.Baselines, tab.ComponentPaneForTests.Page);
            Assert.Equal("BaselinesNavButton", tab.SelectedNavEntryForTests);

            // "2" for the one that differs per region.
            tab.PickComponentForTests("U1");
            await Until(() => tab.ComponentPaneForTests.Label == "U1", "The other component never opened.");

            Assert.Equal("2", Count(tab, "BasicsCountText"));
            Assert.Equal(ComponentPage.Baselines, tab.ComponentPaneForTests.Page);
        });
    }

    // ###########################################################################################
    // THE COMPONENT PICKED STAYS BESIDE A BOARD PART - counted still, and a click on it opens it again,
    // though no selection changes (the list says nothing then).
    // ###########################################################################################
    [Fact]
    public async Task The_component_picked_stays_beside_a_board_part_and_a_click_on_it_opens_it_again()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            tab.PickComponentForTests("U8");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            await tab.ShowEntryAsync(ContributeEntry.Links);

            Assert.Equal("LinksNavButton", tab.SelectedNavEntryForTests);
            Assert.Equal("U8", (list.SelectedItem as ComponentListItem)?.BoardLabel);
            Assert.Equal("1", Count(tab, "PinoutCountText"));

            tab.PickComponentForTests("U8");
            await Until(() => tab.ComponentPaneForTests.BoardPart is null, "The component never opened again.");

            Assert.Equal("BasicsNavButton", tab.SelectedNavEntryForTests);
        });
    }

    [Fact]
    public async Task Another_board_lets_go_of_the_component_picked()
    {
        WriteDraft();
        WriteDraft(BoardB);

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            tab.PickComponentForTests("U8");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            tab.ShowBoard(Entry("250425", BoardB), Board(), "PAL", null);

            Assert.Equal(string.Empty, tab.PickedComponentForTests);
            Assert.Null(list.SelectedItem);
            Assert.Equal("StatusNavButton", tab.SelectedNavEntryForTests);
            Assert.Equal(string.Empty, Count(tab, "BasicsCountText"));
        });
    }

    // "Clear": no component picked, and the page it was on says where to pick one.
    [Fact]
    public async Task Letting_go_of_the_component_keeps_the_page_and_says_where_to_pick_one()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            tab.PickComponentForTests("U8");
            await Until(() => tab.IsComponentOpen, "The component never opened.");
            await tab.ShowEntryAsync(ContributeEntry.ComponentPinout);

            tab.LetGoOfPickedComponent();

            Assert.False(tab.IsComponentOpen);
            Assert.True(tab.IsNoComponentPageShownForTests);
            Assert.Equal("PinoutNavButton", tab.SelectedNavEntryForTests);
            Assert.Null(list.SelectedItem);
        });
    }

    // A rename in the pane: the same component stays picked, as a board part opens beside it.
    [Fact]
    public async Task A_component_renamed_stays_picked()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            tab.PickComponentForTests("C1");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            tab.ComponentPaneForTests.GetControl<TextBox>("BoardLabelBox").Text = "C2";
            tab.ComponentPaneForTests.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("C2", tab.PickedComponentForTests);
        });
    }

    // ###########################################################################################
    // Renamed, and the pane closed on a published board: the column reads the board on screen, which has no C2
    // until it catches up. The pick stays, so C2 is selected once it does (code review, 2026-11-04: the pick
    // went, and its pages said to pick a component). A component the draft does not have either still goes.
    // ###########################################################################################
    [Fact]
    public async Task A_component_renamed_stays_picked_as_the_pane_closes_and_the_board_catches_up()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            tab.PickComponentForTests("C1");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            tab.ComponentPaneForTests.GetControl<TextBox>("BoardLabelBox").Text = "C2";
            tab.ComponentPaneForTests.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            await tab.ShowBoardStatusAsync();

            Assert.False(tab.IsComponentOpen);
            Assert.Equal("C2", tab.PickedComponentForTests);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;

            // The board read again, as Main does once the save is in.
            tab.ShowBoard(Entry(), session.Content.Board, "PAL", null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("C2", tab.PickedComponentForTests);
            Assert.Equal("C2", (list.SelectedItem as ComponentListItem)?.BoardLabel);

            // Gone from the draft and from the board shown: nothing left to pick.
            session.Apply("C2 deleted", content => content with
            {
                Board = content.Board.WithComponentSections(components: content.Board.Components.Where(row => row.BoardLabel != "C2").ToList()),
            });

            tab.ShowBoard(Entry(), session.Content.Board, "PAL", null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(string.Empty, tab.PickedComponentForTests);
        });
    }

    // ---- "UI" or "Table" ------------------------------------------------------------------------------

    private static BoardTableEditor Editor(TabDrafts tab) => tab.GetControl<BoardTableEditor>("TableEditor");

    [Fact]
    public async Task With_Table_picked_each_entry_shows_its_table_and_the_sheet_tabs_are_gone()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.SchematicImages);

            Assert.True(tab.IsTableOpen);
            Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, Editor(tab).CurrentSheet!.Name);
            Assert.Equal("PicturesNavButton", tab.SelectedNavEntryForTests);
            Assert.False(Editor(tab).GetControl<TabControl>("SheetTabs").IsVisible);

            await tab.ShowEntryAsync(ContributeEntry.Credits);

            Assert.Equal(BoardWorkbookSchema.SheetCredits, Editor(tab).CurrentSheet!.Name);
            Assert.Equal("CreditsNavButton", tab.SelectedNavEntryForTests);

            await tab.ShowEntryAsync(ContributeEntry.ComponentReplacements);

            Assert.Equal(BoardWorkbookSchema.SheetComponentReplacements, Editor(tab).CurrentSheet!.Name);
            Assert.Equal("ReplacementsNavButton", tab.SelectedNavEntryForTests);
        });
    }

    // Switching the view keeps what is picked on the left, both ways.
    [Fact]
    public async Task Switching_the_view_keeps_the_entry()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowEntryAsync(ContributeEntry.Links);
            Assert.Equal(PaneBoardPart.Links, tab.ComponentPaneForTests.BoardPart);

            await tab.ShowViewAsync(table: true);

            Assert.True(tab.IsTableOpen);
            Assert.False(tab.IsComponentOpen);
            Assert.Equal(BoardWorkbookSchema.SheetBoardLinks, Editor(tab).CurrentSheet!.Name);
            Assert.Equal("LinksNavButton", tab.SelectedNavEntryForTests);
            Assert.Contains("Selected", tab.GetControl<Button>("TableViewButton").Classes);

            await tab.ShowViewAsync(table: false);

            Assert.False(tab.IsTableOpen);
            Assert.Equal(PaneBoardPart.Links, tab.ComponentPaneForTests.BoardPart);
            Assert.Equal("LinksNavButton", tab.SelectedNavEntryForTests);
            Assert.Contains("Selected", tab.GetControl<Button>("UiViewButton").Classes);
        });
    }

    // ###########################################################################################
    // *** THE ENTRY STAYS ON EVERY KIND OF ENTRY, A BOARD WITH NO DRAFT TOO (owner report, 2026-11-07: "when I am in a
    // menu and I toggle between "UI" and "Table" view, then it should keep the same menu, and not jump to "Overview""). ***
    // A component's page with a component picked, a board part, and an entry that is no sheet - each switched both ways,
    // twice, on a board that had no draft when the switching began.
    // ###########################################################################################
    [Theory]
    [InlineData(nameof(ContributeEntry.Files), "FilesNavButton", false)]
    [InlineData(nameof(ContributeEntry.Files), "FilesNavButton", true)]
    [InlineData(nameof(ContributeEntry.Credits), "CreditsNavButton", false)]
    [InlineData(nameof(ContributeEntry.Regions), "RegionsNavButton", true)]
    [InlineData(nameof(ContributeEntry.Regions), "RegionsNavButton", false)]
    [InlineData(nameof(ContributeEntry.ComponentBasics), "BasicsNavButton", true)]
    [InlineData(nameof(ContributeEntry.ComponentBasics), "BasicsNavButton", false)]
    [InlineData(nameof(ContributeEntry.ComponentPinout), "PinoutNavButton", true)]
    [InlineData(nameof(ContributeEntry.Name), "NameNavButton", true)]
    [InlineData(nameof(ContributeEntry.KiCadFiles), "KiCadNavButton", true)]
    [InlineData(nameof(ContributeEntry.Categories), "CategoriesNavButton", true)]
    [InlineData(nameof(ContributeEntry.ImportantSignals), "SignalsNavButton", true)]
    [InlineData(nameof(ContributeEntry.SchematicImages), "PicturesNavButton", true)]
    [InlineData(nameof(ContributeEntry.ComponentLinks), "ComponentLinksNavButton", true)]
    public async Task Switching_the_view_keeps_the_entry_whatever_it_is(string entryName, string button, bool hasDraft)
    {
        ContributeEntry entry = Enum.Parse<ContributeEntry>(entryName);

        if (hasDraft)
        {
            WriteDraft();
        }

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            // As Main makes the first draft: the published file copied, the board read again and shown.
            tab.EnsureDraftOverrideForTests = board =>
            {
                WriteDraft(board.ExcelDataFile);
                tab.ShowBoard(board, Board(), "PAL", null);
                return Task.FromResult(true);
            };

            if (ContributeEntries.PageOf(entry) is not null)
            {
                tab.PickComponentForTests("U8");
                await Until(() => tab.IsComponentOpen, "The component never opened.");
            }

            await tab.ShowEntryAsync(entry);
            Assert.Equal(button, tab.SelectedNavEntryForTests);

            for (int round = 0; round < 2; round++)
            {
                await tab.ShowViewAsync(table: true);
                Dispatcher.UIThread.RunJobs();
                Assert.Equal(button, tab.SelectedNavEntryForTests);

                await tab.ShowViewAsync(table: false);
                Dispatcher.UIThread.RunJobs();
                Assert.Equal(button, tab.SelectedNavEntryForTests);
            }
        });
    }

    // ###########################################################################################
    // *** "Component images" IS TWO TABLES (owner request, 2026-11-03: "should probably be divided into
    // two"). *** Each its own rows; a picture has no column for what only a baseline has.
    // ###########################################################################################
    [Fact]
    public async Task Pictures_and_Oscilloscope_baseline_are_each_their_own_rows_of_one_sheet()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentPictures);

            DataGrid grid = Editor(tab).GetControl<DataGrid>("TableGrid");

            List<string> Headers() => grid.Columns.Select(column => column.Header?.ToString() ?? string.Empty).Where(header => header.Length > 0).ToList();

            List<BoardTableRow> Rows() => grid.ItemsSource!.Cast<BoardTableRow>().ToList();

            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, Editor(tab).CurrentSheet!.Name);
            Assert.Equal("ComponentPicturesNavButton", tab.SelectedNavEntryForTests);
            Assert.Equal("Top", Assert.Single(Rows()).Cells[3].Text);
            Assert.DoesNotContain(BoardWorkbookSchema.ColPin, Headers());

            await tab.ShowEntryAsync(ContributeEntry.ComponentBaselines);

            Assert.Equal("BaselinesNavButton", tab.SelectedNavEntryForTests);
            Assert.Equal("VSS", Assert.Single(Rows()).Cells[3].Text);
            Assert.Contains(BoardWorkbookSchema.ColPin, Headers());
        });
    }

    // Tab walks the columns shown - walking the sheet's, it stopped at a column the grid does not have.
    [Fact]
    public async Task Tab_moves_along_the_columns_a_part_shows()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentPictures);

            BoardTableEditor editor = Editor(tab);
            BoardTableSheet sheet = editor.CurrentSheet!;
            BoardTableRow row = sheet.Rows.First(candidate => candidate.Cells[3].Text == "Top");

            int region = sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColRegion);
            editor.SelectCell(row, region);
            Assert.Equal(region, editor.CurrentCell!.ColumnIndex);

            editor.MoveToNeighbourCell(forward: true);

            Assert.Equal(BoardWorkbookSchema.ColName, sheet.Columns[editor.CurrentCell!.ColumnIndex]);
        });
    }

    // ###########################################################################################
    // One click on a component is a selection changed and then a tap, each asking for it. With "Table" picked
    // the table is read and built ONCE (code review, 2026-11-04: twice, the guard being only on the pane's way).
    // ###########################################################################################
    [Fact]
    public async Task One_click_on_a_component_opens_the_table_once()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            Assert.False(tab.IsTableOpen);

            tab.PickComponentForTests("U8");
            tab.PickComponentForTests("U8");

            await Until(() => tab.IsTableOpen, "The table never opened.");

            for (int settle = 0; settle < 10; settle++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal(1, tab.TableLoadsForTests);
            Assert.Equal(BoardWorkbookSchema.SheetComponents, Editor(tab).CurrentSheet!.Name);

            // A click later, on the same one, only goes to it: nothing is read again.
            tab.PickComponentForTests("U8");
            await Until(() => Editor(tab).CurrentRow is not null, "The cursor never came to the component.");

            Assert.Equal(1, tab.TableLoadsForTests);
        });
    }

    // ###########################################################################################
    // "Replace" on one part of a sheet in two changes what that part shows, its rows and columns - the
    // baselines and a picture's hidden cells are not on screen, and were rewritten unseen (code review,
    // 2026-11-04).
    // ###########################################################################################
    [Fact]
    public async Task Replace_on_a_part_of_a_sheet_in_two_changes_only_what_that_part_shows()
    {
        BoardData draft = Board();
        draft.ComponentImages[0] = new ComponentImageEntry { BoardLabel = "U8", Name = "Old", File = "Component images/a.png", TimeDiv = "Old" };
        draft.ComponentImages[1] = new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "Old", File = "Scope baseline/b.png" };
        WriteDraft(BoardA, draft);

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentPictures);

            BoardTableEditor editor = Editor(tab);
            BoardTableSheet sheet = editor.CurrentSheet!;

            string Value(string pin, string column)
            {
                List<string> columns = editor.CurrentSheet!.Columns.ToList();

                return editor.CurrentSheet!.Rows
                    .Single(row => !row.IsDeleted && row.Cells[columns.IndexOf(BoardWorkbookSchema.ColPin)].Text == pin)
                    .Cells[columns.IndexOf(column)].Text;
            }

            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "Old";
            editor.GetControl<TextBox>("ReplaceWithBox").Text = "New";
            editor.ReplaceAll();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("New", Value(string.Empty, BoardWorkbookSchema.ColName));
            Assert.Equal("Old", Value(string.Empty, BoardWorkbookSchema.ColTimeDiv));
            Assert.Equal("Old", Value("1", BoardWorkbookSchema.ColName));

            // "On every sheet" went on 2026-11-07 (owner request) - nothing replaces what is not on screen.
            Assert.Null(editor.FindControl<CheckBox>("ReplaceEverySheetBox"));
        });
    }

    // ###########################################################################################
    // "Replace" pressed on a sheet the search narrowed STAYS ON THAT SHEET (owner report, 2026-11-08: "Schematic images"
    // searched for "128", then "Replace" - the table went to "Board: Files" and "Schematic images" was turned off, "Find"
    // still saying "128"). "Replace all" works on the sheet shown alone, so the sheet it opened on is the one it is for.
    // Two things made "Find" find nothing there, neither of them seen: a space after the word, which the search ignores
    // and "Find" looks for as typed, and columns unticked on that sheet earlier under "Columns", kept while CRT runs.
    // ###########################################################################################
    [Theory]
    [InlineData("128 ", false)]
    [InlineData("128", true)]
    public async Task Replace_pressed_on_a_searched_sheet_stays_on_it(string searched, bool columnsUntickedEarlier)
    {
        BoardData draft = Board();
        draft.Schematics[0] = new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "c128-top.png" };
        draft.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Manuals", Name = "C128 manual", File = "Board local files/c128 manual.pdf" });
        WriteDraft(BoardA, draft);

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.SchematicImages);

            BoardTableEditor editor = Editor(tab);

            if (columnsUntickedEarlier)
            {
                editor.OpenReplaceBar();
                editor.UntickEveryReplaceColumn();
                editor.TickReplaceColumn(BoardWorkbookSchema.ColSchematicName, ticked: true);
                editor.GetControl<Button>("CancelReplaceButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                Dispatcher.UIThread.RunJobs();
            }

            editor.SearchText = searched;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, editor.CurrentSheet!.Name);
            Assert.Single(editor.RowsShownForTests, row => !row.IsDeleted);

            editor.OpenReplaceBar();

            for (int settle = 0; settle < 10; settle++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(30);
            }

            // What the table was filtered by, as the search read it - no space after it.
            Assert.Equal("128", editor.GetControl<TextBox>("ReplaceFindBox").Text);
            Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, editor.CurrentSheet!.Name);
            Assert.Equal("PicturesNavButton", tab.SelectedNavEntryForTests);
            Assert.True(tab.GetControl<Button>("PicturesNavButton").IsEnabled);
        });
    }

    // A pill goes to what it found, and the column follows - inside the sheet in two too.
    [Fact]
    public async Task A_pill_moving_the_table_moves_the_entry_selected()
    {
        BoardData draft = Board();
        draft.ComponentImages[1] = new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "Changed", File = "Scope baseline/b.png" };
        WriteDraft(BoardA, draft);

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentPictures);

            Editor(tab).Filter = BoardTableRowKinds.Modified;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardTableSheetPart.Baselines, Editor(tab).SheetPart);
            Assert.Equal("BaselinesNavButton", tab.SelectedNavEntryForTests);

            // What it shows nothing of is off, as such a sheet's tab went; the entry open stays.
            Assert.False(tab.GetControl<Button>("CreditsNavButton").IsEnabled);
            Assert.True(tab.GetControl<Button>("BaselinesNavButton").IsEnabled);

            // An entry with no table is off too: neither a pill nor the search looks in it (owner request, 2026-11-07: "it
            // should also disable the non-table menu entries, as it does not search in those").
            Assert.False(tab.GetControl<Button>("RegionsNavButton").IsEnabled);
            Assert.False(tab.GetControl<Button>("StatusNavButton").IsEnabled);

            Editor(tab).Filter = BoardTableRowKinds.None;
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.GetControl<Button>("CreditsNavButton").IsEnabled);
            Assert.True(tab.GetControl<Button>("RegionsNavButton").IsEnabled);
            Assert.True(tab.GetControl<Button>("StatusNavButton").IsEnabled);

            // The search as the pills: what it cannot look in is off while it is applied.
            Editor(tab).SearchText = "Changed";
            Dispatcher.UIThread.RunJobs();
            Assert.False(tab.GetControl<Button>("RegionsNavButton").IsEnabled);

            Editor(tab).SearchText = string.Empty;
            Dispatcher.UIThread.RunJobs();
            Assert.True(tab.GetControl<Button>("RegionsNavButton").IsEnabled);
        });
    }

    // ###########################################################################################
    // AN ENTRY WITH NO TABLE (five of them: nothing of them is a row) SHOWS ITS PAGE, AND SAYS WHY. The
    // view stays picked, so the next entry with a table shows it.
    // ###########################################################################################
    [Fact]
    public async Task An_entry_with_no_table_shows_its_page_and_says_so()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);

            // "Overview" has none: nothing to open yet.
            Assert.False(tab.IsTableOpen);
            Assert.True(tab.IsNoTableNoticeShownForTests);
            Assert.True(tab.GetControl<StackPanel>("StartPanel").IsVisible);

            await tab.ShowEntryAsync(ContributeEntry.Files);

            Assert.True(tab.IsTableOpen);
            Assert.False(tab.IsNoTableNoticeShownForTests);

            await tab.ShowEntryAsync(ContributeEntry.Regions);

            Assert.False(tab.IsTableOpen);
            Assert.True(tab.IsNoTableNoticeShownForTests);
            Assert.Equal(PaneBoardPart.Regions, tab.ComponentPaneForTests.BoardPart);
            Assert.True(tab.IsTableViewPickedForTests);

            await tab.ShowEntryAsync(ContributeEntry.Links);

            Assert.True(tab.IsTableOpen);
            Assert.Equal(BoardWorkbookSchema.SheetBoardLinks, Editor(tab).CurrentSheet!.Name);
        });
    }

    // The list on the far left goes to the component's rows, opening no pane.
    [Fact]
    public async Task With_Table_picked_a_component_picked_goes_to_its_first_row()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            var window = new Window { Content = new Grid { ColumnDefinitions = new ColumnDefinitions("200,*"), Children = { list, tab } }, Width = 1400, Height = 900 };
            Grid.SetColumn(tab, 1);
            window.Show();
            Dispatcher.UIThread.RunJobs();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);
            Dispatcher.UIThread.RunJobs();

            tab.PickComponentForTests("C1");
            await Until(() => Editor(tab).CurrentRow?.Cells[0].Text == "C1", "The cursor never reached the component's row.");

            Assert.False(tab.IsComponentOpen);
            Assert.True(tab.IsTableOpen);
            Assert.Equal("BasicsNavButton", tab.SelectedNavEntryForTests);

            window.Close();
        });
    }

    // On a sheet with no rows of it the table says so, and stays.
    [Fact]
    public async Task A_component_with_no_rows_on_the_sheet_shown_is_said_to_have_none()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            var window = new Window { Content = new Grid { ColumnDefinitions = new ColumnDefinitions("200,*"), Children = { list, tab } }, Width = 1400, Height = 900 };
            Grid.SetColumn(tab, 1);
            window.Show();
            Dispatcher.UIThread.RunJobs();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentPinout);

            Assert.False(Editor(tab).GoToRowsOfComponent("C1"));
            Assert.Equal(BoardTableEditWording.NoRowsOfComponent("C1"), Editor(tab).GetControl<TextBlock>("StatusText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // WHAT OPENS A PAGE WITH A TABLE OF ITS OWN IS THE "UI" VIEW, however it was asked for: "Edit this
    // component" on the Schematics tab's right-click, with "Table" picked.
    // ###########################################################################################
    [Fact]
    public async Task A_component_opened_from_elsewhere_is_the_UI_view()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.Links);

            await tab.OpenComponentAsync("U8");

            Assert.False(tab.IsTableOpen);
            Assert.True(tab.IsComponentOpen);
            Assert.False(tab.IsTableViewPickedForTests);
            Assert.Contains("Selected", tab.GetControl<Button>("UiViewButton").Classes);
            Assert.False(tab.IsNoTableNoticeShownForTests);
        });
    }

    // "Show in the table" and Submit on a draft with errors say the same the other way round.
    [Fact]
    public async Task A_table_opened_from_elsewhere_is_the_Table_view()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowTableOnSheetAsync(BoardWorkbookSchema.SheetBoardLinks);

            Assert.True(tab.IsTableViewPickedForTests);
            Assert.Contains("Selected", tab.GetControl<Button>("TableViewButton").Classes);
            Assert.Equal("LinksNavButton", tab.SelectedNavEntryForTests);
        });
    }

    // ###########################################################################################
    // THE NEW COMPONENT'S FORM IS NO ENTRY, SO IT HAS NO VIEW TO SWITCH (code review, 2026-11-04). "Table" pressed
    // there took the mode with nothing to show: once added the component's page stood under "Table", and a further
    // press did nothing as the mode already said so. The switch is greyed out while the form is open, and the mode
    // stays what it is if a press lands all the same.
    // ###########################################################################################
    [Fact]
    public async Task The_view_stays_as_it_is_under_the_new_components_form()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();
            ComponentPane pane = tab.ComponentPaneForTests;

            await tab.AddComponentAsync();

            Assert.True(pane.IsNewComponentForm);
            Assert.False(tab.GetControl<Button>("TableViewButton").IsEnabled);

            await tab.ShowViewAsync(table: true);

            Assert.False(tab.IsTableViewPickedForTests);
            Assert.False(tab.IsTableOpen);
            Assert.Contains("Selected", tab.GetControl<Button>("UiViewButton").Classes);

            pane.GetControl<TextBox>("NewLabelBox").Text = "U99";
            pane.GetControl<ComboBox>("NewCategoryBox").SelectedItem = "IC";
            Assert.True(pane.AddNewComponent());

            // Its page, under "UI", and "Table" does what it says.
            Assert.False(tab.IsTableViewPickedForTests);
            Assert.True(tab.GetControl<Button>("TableViewButton").IsEnabled);
            Assert.Equal("BasicsNavButton", tab.SelectedNavEntryForTests);

            await tab.ShowViewAsync(table: true);

            Assert.True(tab.IsTableOpen);
            Assert.True(tab.IsTableViewPickedForTests);
        });
    }

    // Cancel on the form hands the switch back too.
    [Fact]
    public async Task The_view_switch_comes_back_when_the_new_components_form_is_cancelled()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.AddComponentAsync();
            Assert.False(tab.GetControl<Button>("UiViewButton").IsEnabled);

            tab.CloseComponent();

            Assert.True(tab.GetControl<Button>("UiViewButton").IsEnabled);
            Assert.True(tab.GetControl<Button>("TableViewButton").IsEnabled);
        });
    }

    // ###########################################################################################
    // A PICK IN THE LIST WITH "TABLE" ON LEAVES THE KEYS WITH THE LIST (code review, 2026-11-04): the arrow keys
    // walk the components while the table follows. It held from the second pick on, as the table was already
    // open - opening it took the focus whatever was asked, and the next Down moved the grid's cursor instead.
    // An entry picked in the column still hands the keys to the table.
    // ###########################################################################################
    [Fact]
    public async Task A_pick_that_opens_the_table_leaves_the_focus_in_the_list()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = Tab();

            var host = new DockPanel { LastChildFill = true };
            DockPanel.SetDock(list, Avalonia.Controls.Dock.Left);
            host.Children.Add(list);
            host.Children.Add(tab);

            var window = new Window { Content = host, Width = 1300, Height = 800 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            // "Overview" has no table, so nothing is open yet.
            await tab.ShowViewAsync(table: true);
            Assert.False(tab.IsTableOpen);

            // On an item, where the arrow keys are: a ListBox itself takes no focus.
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();
            list.ContainerFromIndex(0)!.Focus();
            Assert.True(list.IsKeyboardFocusWithin);

            tab.PickComponentForTests("U8");
            await Until(() => tab.IsTableOpen, "The table never opened.");

            // The focus is handed over at Background priority, once the table is laid out.
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            Assert.True(list.IsKeyboardFocusWithin, "The list should still have the keys.");
            Assert.False(tab.TableEditor.IsKeyboardFocusWithin);

            // The column's entry is a press for the table.
            await tab.ShowEntryAsync(ContributeEntry.Links);
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.TableEditor.IsKeyboardFocusWithin, "An entry picked in the column should hand the table the keys.");

            window.Close();
        });
    }

    // A draft that could not be made for the table: nothing opens, and what shows stays selected.
    [Fact]
    public async Task An_entry_opens_nothing_when_no_draft_could_be_made_for_its_table()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();
            tab.EnsureDraftOverrideForTests = _ => Task.FromResult(false);

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.Links);

            Assert.False(tab.IsTableOpen);
            Assert.Equal("StatusNavButton", tab.SelectedNavEntryForTests);
        });
    }

    // The clicks reach what they are for - a wrong handler compiles and does nothing.
    [Fact]
    public async Task A_click_on_an_entry_shows_it()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            Click(tab, "CreditsNavButton");
            await Until(() => tab.ComponentPaneForTests.BoardPart == PaneBoardPart.Credits, "The entry never opened.");

            Click(tab, "BasicsNavButton");
            await Until(() => tab.IsNoComponentPageShownForTests, "The page with no component never showed.");
            Assert.Equal("BasicsNavButton", tab.SelectedNavEntryForTests);
        });
    }

    [Fact]
    public async Task A_click_on_a_view_shows_it()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowEntryAsync(ContributeEntry.Files);

            tab.GetControl<Button>("TableViewButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsTableOpen, "The table never opened.");

            tab.GetControl<Button>("UiViewButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => !tab.IsTableOpen, "The table never closed.");
        });
    }

    // ---- several components picked, edited as one (owner request, 2026-11-08) ---------------------------

    // The tab with the list on the far left as Main has it: several can be picked, with Ctrl or Shift.
    private static (TabDrafts Tab, ListBox List) TabPickingSeveral()
    {
        (TabDrafts tab, ListBox list) = Tab();
        list.SelectionMode = SelectionMode.Multiple;

        return (tab, list);
    }

    [Fact]
    public async Task Several_picked_open_together_in_the_pane_and_one_picked_again_is_that_ones()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = TabPickingSeveral();

            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);
            tab.PickComponentsForTests("U8", "C1");

            await Until(() => tab.ComponentPaneForTests.IsSeveral, "The two never opened together.");

            Assert.Equal(["U8", "C1"], tab.ComponentPaneForTests.Labels);
            Assert.Equal(["U8", "C1"], tab.PickedComponentsForTests);
            Assert.Equal(2, list.SelectedItems!.Count);
            Assert.Equal("BasicsNavButton", tab.SelectedNavEntryForTests);

            // One picked again: the pane is that one's.
            tab.PickComponentForTests("C1");
            await Until(() => !tab.ComponentPaneForTests.IsSeveral, "The pane stayed on the two.");

            Assert.Equal(["C1"], tab.PickedComponentsForTests);
        });
    }

    // ###########################################################################################
    // A PICK MADE WHILE THE ONE BEFORE IS STILL OPENING STANDS (seen as a Shift+click failing now and then, 2026-11-08): the
    // first open, ending last, put the pane on the one component and the pick back to it - the list with it.
    // ###########################################################################################
    [Fact]
    public async Task Several_picked_while_one_is_still_opening_stay_picked()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = TabPickingSeveral();

            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);

            // The first open held until the second pick has opened - so it ends last, as it can under load.
            var hold = new TaskCompletionSource();
            tab.HoldOpeningForTests = hold.Task;
            tab.PickComponentForTests("U8");
            tab.HoldOpeningForTests = null;

            tab.PickComponentsForTests("U8", "C1");
            await Until(() => tab.ComponentPaneForTests.IsSeveral, "The two never opened together.");

            hold.SetResult();

            for (int settle = 0; settle < 30; settle++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal(["U8", "C1"], tab.PickedComponentsForTests);
            Assert.True(tab.ComponentPaneForTests.IsSeveral);
            Assert.Equal(["U8", "C1"], tab.ComponentPaneForTests.Labels);
            Assert.Equal(2, list.SelectedItems!.Count);
        });
    }

    // ###########################################################################################
    // A CLICK ON THE ONE STILL OPENING, AFTER SEVERAL WERE PICKED, IS THE PICK (code review, 2026-11-09): it was dropped as
    // "already opening", and the several - ending last - took the pick back.
    // ###########################################################################################
    [Fact]
    public async Task A_click_on_the_one_still_opening_after_several_were_picked_is_the_pick()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = TabPickingSeveral();

            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);

            var hold = new TaskCompletionSource();
            tab.HoldOpeningForTests = hold.Task;
            tab.PickComponentForTests("U8");
            tab.HoldOpeningForTests = null;

            tab.PickComponentsForTests("U8", "C1");
            await Until(() => tab.ComponentPaneForTests.IsSeveral, "The two never opened together.");

            tab.PickComponentsForTests("U8");
            hold.SetResult();

            await Until(() => tab.IsComponentOpen && !tab.ComponentPaneForTests.IsSeveral, "U8 alone never opened.");

            for (int settle = 0; settle < 30; settle++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal(["U8"], tab.PickedComponentsForTests);
            Assert.Equal(["U8"], tab.ComponentPaneForTests.Labels);
            Assert.Single(list.SelectedItems!);
        });
    }

    // ###########################################################################################
    // "MARK ALL" ON THE OTHER TABS IS SHOWN SELECTED HERE, NEVER OPENED FOR EDITING (code review, 2026-11-09): taken as the
    // tab was entered, it opened the pane on every component, and the next field typed went to all of them. Fewer than all
    // open together, as before.
    // ###########################################################################################
    [Fact]
    public async Task Mark_all_taken_from_the_other_tabs_is_shown_selected_and_never_opened_for_editing()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, ListBox list) = TabPickingSeveral();

            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);
            tab.PickComponentForTests("U8");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            List<string> every = tab.ComponentItems.Select(item => item.BoardLabel).ToList();
            Assert.True(every.Count > 2);

            tab.TakePickFromOtherTabs(every, everyRowOfTheirs: true);
            Dispatcher.UIThread.RunJobs();

            Assert.False(tab.IsComponentOpen);
            Assert.True(tab.IsNoComponentPageShownForTests);
            Assert.Contains("Every component in the list is selected", tab.NoComponentTextForTests, StringComparison.Ordinal);
            Assert.Equal(every, tab.PickedComponentsForTests);
            Assert.Equal(every.Count, list.SelectedItems!.Count);

            // Nor does a page picked afterwards open them, nor "Table" and back (second code review, 2026-11-09: a page
            // picked afterwards opened every component of the board as one).
            await tab.ShowEntryAsync(ContributeEntry.ComponentPinout);
            Assert.False(tab.IsComponentOpen);
            Assert.Contains("Every component in the list is selected", tab.NoComponentTextForTests, StringComparison.Ordinal);
            Assert.Equal("PinoutNavButton", tab.SelectedNavEntryForTests);

            await tab.ShowViewAsync(table: true);
            await Until(() => tab.IsTableOpen, "The table never opened.");

            // The table every row, as with none picked (second code review, 2026-11-09: narrowed to the labels both lists had).
            Assert.Empty(Editor(tab).ComponentsShownForTests);
            Assert.Contains("RowsDraggable", Editor(tab).GetControl<DataGrid>("TableGrid").Classes);

            await tab.ShowViewAsync(table: false);
            await Until(() => !tab.IsTableOpen, "The table never closed.");
            Assert.False(tab.IsComponentOpen);
            Assert.Equal(every, tab.PickedComponentsForTests);

            // Not on a component's page as it is taken: the same once one is picked.
            await tab.ShowEntryAsync(ContributeEntry.Overview);
            tab.TakePickFromOtherTabs(["U8"]);
            tab.TakePickFromOtherTabs(every, everyRowOfTheirs: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);
            Assert.False(tab.IsComponentOpen);

            // One clicked in the list is the pick, opened as ever.
            tab.PickComponentsForTests("C1");
            await Until(() => tab.IsComponentOpen, "C1 never opened.");
            Assert.Equal(["C1"], tab.ComponentPaneForTests.Labels);

            tab.TakePickFromOtherTabs(["U8", "C1"]);
            await Until(() => tab.ComponentPaneForTests.IsSeveral, "The two never opened together.");
            Assert.Equal(["U8", "C1"], tab.ComponentPaneForTests.Labels);
        });
    }

    // In "Table" the component sheets show the rows of the components picked alone - and "Replace" replaces in those.
    [Fact]
    public async Task In_the_table_the_component_sheets_show_the_rows_of_the_components_picked_and_replace_only_those()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = TabPickingSeveral();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);
            await Until(() => tab.IsTableOpen, "The table never opened.");

            BoardTableEditor editor = Editor(tab);
            int label = editor.CurrentSheet!.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel);
            List<string> Shown() => editor.RowsShownForTests.Select(row => row.Cells[label].Text).ToList();

            Assert.Equal(4, Shown().Count);

            tab.PickComponentsForTests("U8", "C1");
            await Until(() => editor.ComponentsShownForTests.Count == 2, "The table never narrowed.");

            Assert.Equal(["U8", "C1"], Shown());
            Assert.True(editor.GetControl<TextBlock>("ComponentsShownText").IsVisible);

            // Narrowed, a component sheet's rows do not move by their grip.
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");
            Assert.DoesNotContain("RowsDraggable", grid.Classes);

            int friendly = editor.CurrentSheet!.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "i";
            editor.GetControl<TextBox>("ReplaceWithBox").Text = "I";
            editor.ReplaceAll();
            Dispatcher.UIThread.RunJobs();

            // "Filter" of C1 replaced, "VIC" of U1 - not picked - left alone.
            Assert.Equal("FIlter", editor.CurrentSheet!.Rows.Single(row => row.Cells[label].Text == "C1").Cells[friendly].Text);
            Assert.All(editor.CurrentSheet!.Rows.Where(row => row.Cells[label].Text == "U1"), row => Assert.Equal("VIC", row.Cells[friendly].Text));

            // Another entry picked while the bar is open is the sheet shown, though "Find" finds nothing there (owner report,
            // 2026-11-08: the table went back to a sheet "Find" had something on).
            await tab.ShowEntryAsync(ContributeEntry.Links);
            Assert.Equal(BoardWorkbookSchema.SheetBoardLinks, editor.CurrentSheet!.Name);
            Assert.Equal("LinksNavButton", tab.SelectedNavEntryForTests);

            // A sheet of the board as a whole is not narrowed.
            editor.GetControl<Button>("CancelReplaceButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();
            Assert.Single(editor.RowsShownForTests);
            Assert.False(editor.GetControl<TextBlock>("ComponentsShownText").IsVisible);

            // ... so its rows move by their grip whatever components are picked (code review, 2026-11-09: they did not).
            Assert.Equal(2, editor.ComponentsShownForTests.Count);
            Assert.Contains("RowsDraggable", grid.Classes);
        });
    }

    // ###########################################################################################
    // A PILL WITH COMPONENTS PICKED GOES BY THE ROWS THE TABLE SHOWS (third code review, 2026-11-09): "Errors" counted over the
    // draft left "Components: Links" on for C1's error while the table showed U8's rows alone - picked, it showed nothing.
    // ###########################################################################################
    [Fact]
    public async Task A_pill_with_a_component_picked_turns_off_the_entries_whose_rows_the_pick_hides()
    {
        // Errors on two sheets: U8 with no category (Basics), C1's link (Links).
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", Category = "", FriendlyName = "CPU" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", FriendlyName = "Filter" },
            ],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "C1", Name = "Bad", Url = "ftp://example.com" }],
        };

        WriteDraft(board: board);

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);
            await Until(() => tab.IsTableOpen, "The table never opened.");

            BoardTableEditor editor = Editor(tab);
            editor.Filter = BoardTableRowKinds.Errors;
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(BoardWorkbookSchema.SheetComponents, editor.CurrentSheet!.Name);
            Assert.True(tab.GetControl<Button>("ComponentLinksNavButton").IsEnabled);

            tab.PickComponentsForTests("U8");
            await Until(() => editor.ComponentsShownForTests.Count == 1, "The table never narrowed.");
            Dispatcher.UIThread.RunJobs();

            Assert.False(tab.GetControl<Button>("ComponentLinksNavButton").IsEnabled);

            // C1 picked: on again.
            tab.PickComponentsForTests("C1");
            await Until(() => editor.ComponentsShownForTests.Contains("C1"), "The table never narrowed to C1.");
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.GetControl<Button>("ComponentLinksNavButton").IsEnabled);
        });
    }

    // ###########################################################################################
    // *** SUBMIT ON A DRAFT WITH ERRORS GOES TO THE FIRST IN THE VIEW PICKED (owner request, 2026-11-13: "no matter the view
    // I am in, then it will go to the "Table" view and show the error cells. Instead I think it should use the current view
    // as-is, and then navigate to the problem"). *** With U8 picked and the error on C1's link: in "UI" C1 opens on its
    // "Links" page, whose panel lists the error, and the tab's line says why nothing was sent - no table opened.
    // ###########################################################################################
    private static BoardData BoardWithAnErrorOnC1() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" },
            new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", FriendlyName = "Filter" },
        ],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "C1", Name = "Bad", Url = "ftp://example.com" }],
    };

    [Fact]
    public async Task Submit_blocked_by_errors_in_the_UI_view_opens_the_first_on_its_page()
    {
        BoardData board = BoardWithAnErrorOnC1();
        WriteDraft(board: board);

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            tab.PickComponentForTests("U8");
            await Until(() => tab.IsComponentOpen, "The component never opened.");

            Assert.True(await tab.StopForErrorsAsync(Entry(), board));
            await Until(() => tab.EntryOpenForTests == ContributeEntry.ComponentLinks && tab.PickedComponentsForTests.SequenceEqual(["C1"]), "C1's links never opened.");

            Assert.False(tab.IsTableOpen);
            Assert.Equal("ComponentLinksNavButton", tab.SelectedNavEntryForTests);
            Assert.Contains(tab.EntryProblemRowsForTests, row => row.IsError && row.Component == "C1");
            Assert.True(tab.IsEntryProblemsPanelShownForTests);
            Assert.StartsWith("This draft has 1 error to fix", tab.DraftProblemForTests, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // The same in "Table": the sheet of the first error, narrowed to its component as a pick narrows it, the cursor on its
    // cell, and the table's line saying so. Nothing is narrowed to the errors any more - the "Errors" pill is the user's (it
    // was picked by Submit until 2026-11-13, which is why a picked component had to be let go of: code review, 2026-11-09).
    // ###########################################################################################
    [Fact]
    public async Task Submit_blocked_by_errors_in_the_Table_view_puts_the_cursor_on_the_first()
    {
        BoardData board = BoardWithAnErrorOnC1();
        WriteDraft(board: board);

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = Tab();

            await tab.ShowViewAsync(table: true);
            tab.PickComponentForTests("U8");
            await Until(() => Editor(tab).ComponentsShownForTests.Contains("U8"), "The table never narrowed to U8.");

            Assert.True(await tab.StopForErrorsAsync(Entry(), board));

            BoardTableEditor editor = Editor(tab);
            await Until(() => editor.CurrentSheet?.Name == BoardWorkbookSchema.SheetComponentLinks && editor.ComponentsShownForTests.Contains("C1"), "The table never went to C1's links.");

            Assert.True(tab.IsTableOpen);
            Assert.Equal(BoardTableRowKinds.None, editor.Filter);
            Assert.Equal(["C1"], tab.PickedComponentsForTests);

            BoardTableSheet sheet = editor.CurrentSheet!;
            Assert.Equal(BoardWorkbookSchema.ColUrl, sheet.Columns[editor.CurrentCell!.ColumnIndex]);
            Assert.Equal("C1", editor.CurrentRow!.Cells[sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel)].Text);
            Assert.StartsWith("This draft has 1 error to fix before it can be submitted. The cursor is on it.", editor.GetControl<TextBlock>("StatusText").Text, StringComparison.Ordinal);
            Assert.Null(tab.DraftProblemForTests);
        });
    }
}
