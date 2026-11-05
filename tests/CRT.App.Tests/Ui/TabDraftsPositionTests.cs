using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// BACK ON THE TAB, EVERYTHING STANDS WHERE IT STOOD (owner request, 2026-11-04: "navigate to the exact same
// position in the 'Contribute' tab, if toggling between the main tabs (e.g. going to 'Schematics' tab and
// then back again), as I expect there will be a lot of preview in real app").
//
// A tab not shown is only detached, so what the tab holds keeps its place by itself - measured that day, and
// held here, as anything reading it again on the way back would lose it. Through a real TabControl, the
// board read again as it is left, as Main does once the tab's saves left the board on screen behind. The lists
// on the far left are Main's: MainContributeListsTests holds them.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsPositionTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsPositionTests() => DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static BoardData LongBoard() => new()
    {
        RevisionDate = "2026-October-4",
        Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
        Components = [.. Enumerable.Range(1, 40).Select(i => new ComponentEntry { BoardLabel = $"U{i}", Category = "IC", FriendlyName = $"Part {i}" })],
        ComponentPins = [.. Enumerable.Range(1, 80).Select(i => new ComponentPinEntry { BoardLabel = "U1", Pin = $"{i}", Name = $"P{i}", Direction = "Input" })],
        BoardLinks = [.. Enumerable.Range(1, 80).Select(i => new BoardLinkEntry { Category = "Docs", Name = $"Link {i}", Url = $"https://example.com/{i}" })],
    };

    private static void Pump()
    {
        for (int i = 0; i < 6; i++)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }
    }

    // The tab in a real TabControl, the board read again as it asks - as Main reads it.
    private sealed record Host(TabDrafts Tab, TabControl Tabs, TabItem Contribute, TabItem Other, Window Window)
    {
        public void GoAwayAndBack()
        {
            this.Tabs.SelectedItem = this.Other;
            Pump();
            this.Tabs.SelectedItem = this.Contribute;
            Pump();
        }
    }

    private static Host Hosted()
    {
        var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = BoardKey };
        var tab = new TabDrafts { HardwareBoardsOverrideForTests = [entry], PublishedBoardOverrideForTests = _ => null };

        tab.ReloadBoardOverrideForTests = _ => tab.ShowBoard(entry, LongBoard(), "PAL", null);
        tab.RefreshDrafts();

        var contribute = new TabItem { Header = "Contribute", Content = tab };
        var other = new TabItem { Header = "Schematics", Content = new TextBlock { Text = "Elsewhere" } };
        var tabs = new TabControl { ItemsSource = new[] { contribute, other } };
        var window = new Window { Width = 1300, Height = 700, Content = tabs };

        window.Show();
        Pump();

        tab.ShowBoard(entry, LongBoard(), "PAL", null);
        Pump();

        return new Host(tab, tabs, contribute, other, window);
    }

    private void WriteDraft()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardKey);

        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, LongBoard());
        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, BoardKey), new DraftMarker { BoardKey = BoardKey });
    }

    private static ScrollViewer LongScroller(Control within) =>
        within.GetVisualDescendants().OfType<ScrollViewer>().First(viewer => viewer.IsEffectivelyVisible && viewer.Extent.Height > viewer.Viewport.Height + 50);

    [Fact]
    public async Task A_component_stays_on_its_page_scrolled_where_it_was_with_a_change_saved_on_the_way_out()
    {
        this.WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            Host host = Hosted();
            ComponentPane pane = host.Tab.ComponentPaneForTests;

            await host.Tab.OpenComponentAsync("U1");
            pane.ShowPage(ComponentPage.Pinout);
            Pump();

            ScrollViewer sections = pane.GetControl<ScrollViewer>("SectionsScrollViewer");
            sections.Offset = new Vector(0, 700);
            Pump();

            // A change waiting: it goes as the tab is left, and the board is read again.
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "DIP-40";
            pane.CommitHeld();

            host.GoAwayAndBack();
            await pane.LastSave;
            Pump();

            Assert.True(host.Tab.IsComponentOpen);
            Assert.Equal("U1", pane.Label);
            Assert.Equal(ComponentPage.Pinout, pane.Page);
            Assert.Equal(ContributeEntry.ComponentPinout, host.Tab.EntryOpenForTests);
            Assert.Equal(700, sections.Offset.Y, 1);

            host.Window.Close();
        });
    }

    [Fact]
    public async Task A_part_of_the_board_stays_open_scrolled_where_it_was()
    {
        this.WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            Host host = Hosted();

            await host.Tab.OpenBoardPartAsync(PaneBoardPart.Links);
            Pump();

            ScrollViewer links = LongScroller(host.Tab.ComponentPaneForTests);
            links.Offset = new Vector(0, 600);
            Pump();

            host.GoAwayAndBack();

            Assert.Equal(ContributeEntry.Links, host.Tab.EntryOpenForTests);
            Assert.Same(links, LongScroller(host.Tab.ComponentPaneForTests));
            Assert.Equal(600, links.Offset.Y, 1);

            host.Window.Close();
        });
    }

    [Fact]
    public async Task The_table_stays_on_its_sheet_and_its_rows_with_the_view_picked()
    {
        this.WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            Host host = Hosted();

            await host.Tab.ShowViewAsync(table: true);
            await host.Tab.ShowEntryAsync(ContributeEntry.Links);
            Pump();

            BoardTableEditor editor = host.Tab.GetControl<BoardTableEditor>("TableEditor");
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            grid.ScrollIntoView(grid.ItemsSource!.Cast<object>().Last(), null);
            Pump();

            ScrollBar bar = grid.GetVisualDescendants().OfType<ScrollBar>().First(scroll => scroll.Orientation == Avalonia.Layout.Orientation.Vertical);
            double scrolled = bar.Value;
            string? sheet = editor.CurrentSheet?.Name;

            Assert.True(scrolled > 0);

            host.GoAwayAndBack();

            Assert.True(host.Tab.IsTableViewPickedForTests);
            Assert.Equal(ContributeEntry.Links, host.Tab.EntryOpenForTests);
            Assert.Equal(sheet, editor.CurrentSheet?.Name);
            Assert.Equal(scrolled, grid.GetVisualDescendants().OfType<ScrollBar>().First(scroll => scroll.Orientation == Avalonia.Layout.Orientation.Vertical).Value, 1);

            host.Window.Close();
        });
    }

    [Fact]
    public void The_left_column_stays_scrolled_on_a_low_window()
    {
        this.WriteDraft();

        UiTest.Run(() =>
        {
            Host host = Hosted();

            host.Window.Height = 400;
            Pump();

            ScrollViewer nav = host.Tab.GetControl<ScrollViewer>("BoardNavScroll");
            nav.Offset = new Vector(0, 120);
            Pump();

            host.GoAwayAndBack();

            Assert.Equal(120, nav.Offset.Y, 1);

            host.Window.Close();
        });
    }
}
