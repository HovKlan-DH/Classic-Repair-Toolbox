using Avalonia;
using Avalonia.Controls;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE DRAFT'S ERRORS AND WARNINGS, WHERE THEY ARE (TabDrafts.Problems.cs - owner requests, 2026-11-04: "move
// the error and warning badges to where the problem is"; "the 'Overview' entry should show all errors and
// warnings, and then you can see the details per entry in the list below"; and with a component selected "I
// do not want to see any warnings or errors for other components").
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsProblemsTests : IDisposable
{
    private const string BoardA = "Commodore/C64/250407/board.json";
    private const string BoardB = "Commodore/C64/250425/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsProblemsTests() => DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry Entry(string board = "250407", string excelDataFile = BoardA) =>
        new() { HardwareName = "Commodore 64", BoardName = board, ExcelDataFile = excelDataFile };

    // Every component marked, so what is left is what each test is about: no category on U9, a bad address
    // on U8's link and U9's, a credit with nobody named, a warning on each of two board links, and an image
    // with no file. C1 has nothing wrong.
    private static BoardData Board() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC" },
            new ComponentEntry { BoardLabel = "U9", FriendlyName = "VIC" },
            new ComponentEntry { BoardLabel = "C1", FriendlyName = "Filter", Category = "Capacitor" },
        ],
        ComponentHighlights =
        [
            new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "1", Y = "1", Width = "5", Height = "5" },
            new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U9", X = "1", Y = "1", Width = "5", Height = "5" },
            new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "C1", X = "1", Y = "1", Width = "5", Height = "5" },
        ],
        ComponentLinks =
        [
            new ComponentLinkEntry { BoardLabel = "U8", Name = "Bad one", Url = "ftp://example.com/u8" },
            new ComponentLinkEntry { BoardLabel = "U9", Name = "Bad too", Url = "ftp://example.com/u9" },
        ],
        Credits = [new CreditEntry { Category = "Schematic image", SubCategory = "Top" }],
        BoardLinks =
        [
            new BoardLinkEntry { Category = "Web", Name = "Same", Url = "https://example.com/a" },
            new BoardLinkEntry { Category = "Web", Name = "Same", Url = "https://example.com/a" },
        ],
    };

    private static void WriteDraft(BoardData board, string excelDataFile = BoardA)
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker { BoardKey = excelDataFile });
    }

    private static TabDrafts Tab(BoardData board)
    {
        HardwareBoardEntry entry = Entry();

        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [entry, Entry("250425", BoardB)],
            PublishedBoardOverrideForTests = _ => null,
        };

        tab.RefreshDrafts();
        tab.ShowBoard(entry, board, "PAL", null);

        return tab;
    }

    private static string Title(TabDrafts tab) => tab.GetControl<TextBlock>("EntryProblemsTitle").Text ?? string.Empty;

    private static string Help(TabDrafts tab) => tab.GetControl<TextBlock>("EntryProblemsHelp").Text ?? string.Empty;

    private static ComponentListItem Row(TabDrafts tab, string label) => tab.ComponentItems.Single(item => item.BoardLabel == label);

    // ###########################################################################################
    // "OVERVIEW" HAS THEM ALL, AND THE ENTRIES BELOW IT WHERE THEY ARE - which add up to the same.
    // ###########################################################################################
    [Fact]
    public void Overview_counts_every_problem_and_the_entries_below_it_add_up_to_the_same()
    {
        BoardData board = Board();
        WriteDraft(board);

        UiTest.Run(() =>
        {
            TabDrafts tab = Tab(board);
            DraftListItem row = Assert.Single(tab.Drafts);

            Assert.Equal(new BoardProblemCounts(5, 2), row.Problems);
            Assert.Equal(("5", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));

            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentBasics));
            Assert.Equal(("2", string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentLinks));
            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.Credits));
            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.SchematicImages));
            Assert.Equal((string.Empty, "2"), tab.NavBadgesForTests(ContributeEntry.Links));

            // Nothing on what has nothing wrong.
            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.Files));
            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentPinout));

            int Sum(Func<(string Errors, string Warnings), string> of) =>
                Enum.GetValues<ContributeEntry>()
                    .Where(entry => entry != ContributeEntry.Overview)
                    .Select(entry => of(tab.NavBadgesForTests(entry)))
                    .Sum(text => text.Length == 0 ? 0 : int.Parse(text, System.Globalization.CultureInfo.InvariantCulture));

            Assert.Equal(row.Problems.Errors, Sum(badges => badges.Errors));
            Assert.Equal(row.Problems.Warnings, Sum(badges => badges.Warnings));

            // No panel saying so over the page (owner request, 2026-11-04: "there is no need to show the first
            // panel ... The user can scroll down and see exactly where the issues are"): its panels say each.
            Assert.False(tab.IsEntryProblemsPanelShownForTests);
            Assert.Empty(tab.EntryProblemRowsForTests);
        });
    }

    // Another board has its own badges - and one with nothing published to look at, none.
    [Fact]
    public void Another_board_shown_has_its_own_badges_and_one_with_no_draft_and_nothing_published_none()
    {
        BoardData board = Board();
        WriteDraft(board);

        UiTest.Run(() =>
        {
            TabDrafts tab = Tab(board);
            tab.PublishedBoardFileOverrideForTests = _ => Path.Combine(Path.GetTempPath(), "crt-no-published-board-" + Guid.NewGuid().ToString("N"), "board.json");
            Assert.Equal(("5", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));

            tab.ShowBoard(Entry("250425", BoardB), new BoardData(), "PAL", null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.Overview));
            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentBasics));
            Assert.False(tab.IsEntryProblemsPanelShownForTests);
        });
    }

    // ###########################################################################################
    // *** A PUBLISHED BOARD WITH NO DRAFT SAYS ITS OWN (owner decision, 2026-11-07: the launch check of every board,
    // written to the log alone, "needs to run for everyone and its findings should show just like any other error or
    // warning in the "Contribute" tab"). *** The same badges a draft of it would have, the same rows on the component
    // list; a draft made of it then has its own, and the board going back to no draft its published board's again.
    // ###########################################################################################
    [Fact]
    public async Task A_published_board_with_no_draft_shows_its_own_errors_and_warnings_as_a_draft_would()
    {
        string publishedFile = this.thisWorkspace.Path_("Data", "Commodore", "C64", "250425", "board.json");
        Directory.CreateDirectory(Path.GetDirectoryName(publishedFile)!);
        CachedWorkbooks.Write(publishedFile, Board());

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(new BoardData());
            tab.PublishedBoardFileOverrideForTests = entry => entry.ExcelDataFile == BoardB ? publishedFile : Path.Combine(Path.GetTempPath(), "crt-none-" + Guid.NewGuid().ToString("N"));

            tab.ShowBoard(Entry("250425", BoardB), Board(), "PAL", null);

            for (int attempt = 0; attempt < 300 && tab.NavBadgesForTests(ContributeEntry.Overview) == (string.Empty, string.Empty); attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal(("5", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));
            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentBasics));
            Assert.Equal((string.Empty, "2"), tab.NavBadgesForTests(ContributeEntry.Links));
            Assert.True(Row(tab, "U9").Problems.Errors > 0);
            Assert.Equal(new BoardProblemCounts(0, 0), Row(tab, "C1").Problems);

            // Another board shown lets them go.
            tab.ShowBoard(Entry(), new BoardData(), "PAL", null);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.Overview));
        });
    }

    // ###########################################################################################
    // *** A PUBLISHED BOARD IS CHECKED ONLY ONCE THIS TAB IS SHOWN (code review, 2026-11-08). *** Every board shown -
    // on the "Schematics" tab, switching back and forth - started the published board's whole check, every row and
    // every file it cites, for badges only this tab shows, and only one board's read is kept, so each switch checked
    // the board again. Off screen the tab checks nothing; shown, it checks the board on screen and the badges come.
    // ###########################################################################################
    [Fact]
    public async Task A_published_board_is_checked_only_once_the_tab_is_shown()
    {
        string publishedFile = this.thisWorkspace.Path_("Data", "Commodore", "C64", "250425", "board.json");
        Directory.CreateDirectory(Path.GetDirectoryName(publishedFile)!);
        CachedWorkbooks.Write(publishedFile, Board());

        await UiTest.RunAsync(async () =>
        {
            bool onScreen = false;

            var tab = new TabDrafts
            {
                HardwareBoardsOverrideForTests = [Entry(), Entry("250425", BoardB)],
                PublishedBoardOverrideForTests = _ => null,
                PublishedBoardFileOverrideForTests = entry => entry.ExcelDataFile == BoardB ? publishedFile : Path.Combine(Path.GetTempPath(), "crt-none-" + Guid.NewGuid().ToString("N")),
                IsOnScreenOverrideForTests = () => onScreen,
            };

            tab.RefreshDrafts();
            tab.ShowBoard(Entry("250425", BoardB), Board(), "PAL", null);

            // Time enough for a check of this small board to have come back, had one started.
            for (int attempt = 0; attempt < 30; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.Overview));
            Assert.Equal(0, tab.PublishedChecksForTests);

            onScreen = true;
            var window = new Window { Content = tab, Width = 1200, Height = 800 };
            window.Show();

            try
            {
                for (int attempt = 0; attempt < 300 && tab.NavBadgesForTests(ContributeEntry.Overview) == (string.Empty, string.Empty); attempt++)
                {
                    Dispatcher.UIThread.RunJobs();
                    await Task.Delay(10);
                }

                Assert.Equal(("5", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));
                Assert.True(tab.PublishedChecksForTests >= 1);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // *** THE PUBLISHED BOARD SHOWN IS CHECKED AGAIN ONCE A DATA SYNC IS OVER (code review, 2026-11-08). *** The files a
    // board file names sync after CRT's window opens, so a published board checked meanwhile said "is not in the
    // downloaded data" of files still on their way, until the board or the tab was shown again. Asked again, the badges'
    // check and the board status's panels both run once more (what they find then is DraftStatusReader's, which keeps a
    // read only while every file it missed is still missing - DraftStatusReaderTests).
    // ###########################################################################################
    [Fact]
    public async Task The_published_board_shown_is_checked_again_after_a_data_sync()
    {
        string publishedFile = this.thisWorkspace.Path_("Data", "Commodore", "C64", "250425", "board.json");
        Directory.CreateDirectory(Path.GetDirectoryName(publishedFile)!);
        CachedWorkbooks.Write(publishedFile, Board());

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(new BoardData());
            tab.PublishedBoardFileOverrideForTests = entry => entry.ExcelDataFile == BoardB ? publishedFile : Path.Combine(Path.GetTempPath(), "crt-none-" + Guid.NewGuid().ToString("N"));

            tab.ShowBoard(Entry("250425", BoardB), Board(), "PAL", null);

            for (int attempt = 0; attempt < 300 && tab.NavBadgesForTests(ContributeEntry.Overview) == (string.Empty, string.Empty); attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal(("5", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));
            int checks = tab.PublishedChecksForTests;
            int asks = tab.ReadinessAsksForTests;

            tab.CheckPublishedBoardAgain();

            for (int attempt = 0; attempt < 300 && tab.PublishedChecksForTests == checks; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.True(tab.PublishedChecksForTests > checks, "The published board was not checked again in time");
            Assert.True(tab.ReadinessAsksForTests > asks);
            Assert.Equal(("5", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));
        });
    }

    // ###########################################################################################
    // *** A CHECK ASKED FOR AFTER A DATA SYNC IS A NEW ONE, NEVER ONE STILL RUNNING FROM BEFORE IT (code review, 2026-11-08).
    // *** The check still running for the board began before the sync's files arrived, and was handed back as the check asked
    // for after it: its "not in the downloaded data" stayed until the board was shown again. Held here at its end, it is
    // still running as the sync's check is asked for.
    // ###########################################################################################
    [Fact]
    public async Task A_check_asked_for_after_a_data_sync_never_reuses_one_still_running()
    {
        string publishedFile = this.thisWorkspace.Path_("Data", "Commodore", "C64", "250425", "board.json");
        Directory.CreateDirectory(Path.GetDirectoryName(publishedFile)!);
        CachedWorkbooks.Write(publishedFile, Board());

        using var held = new ManualResetEventSlim();
        DraftStatusReader.BeforePublishedReadIsKeptForTests = file =>
        {
            if (string.Equals(Path.GetFullPath(file), Path.GetFullPath(publishedFile), StringComparison.OrdinalIgnoreCase))
            {
                held.Wait(TimeSpan.FromSeconds(10));
            }
        };

        try
        {
            await UiTest.RunAsync(async () =>
            {
                TabDrafts tab = Tab(new BoardData());
                tab.PublishedBoardFileOverrideForTests = entry => entry.ExcelDataFile == BoardB ? publishedFile : Path.Combine(Path.GetTempPath(), "crt-none-" + Guid.NewGuid().ToString("N"));

                tab.ShowBoard(Entry("250425", BoardB), Board(), "PAL", null);

                for (int attempt = 0; attempt < 300 && tab.PublishedChecksForTests == 0; attempt++)
                {
                    Dispatcher.UIThread.RunJobs();
                    await Task.Delay(10);
                }

                int checks = tab.PublishedChecksForTests;
                Assert.True(checks > 0, "The published board's check did not start in time");

                tab.CheckPublishedBoardAgain();

                for (int attempt = 0; attempt < 300 && tab.PublishedChecksForTests == checks; attempt++)
                {
                    Dispatcher.UIThread.RunJobs();
                    await Task.Delay(10);
                }

                Assert.True(tab.PublishedChecksForTests > checks, "The check still running was handed back for the sync's");

                held.Set();

                for (int attempt = 0; attempt < 300 && tab.NavBadgesForTests(ContributeEntry.Overview) == (string.Empty, string.Empty); attempt++)
                {
                    Dispatcher.UIThread.RunJobs();
                    await Task.Delay(10);
                }

                Assert.Equal(("5", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));
            });
        }
        finally
        {
            held.Set();
            DraftStatusReader.BeforePublishedReadIsKeptForTests = null;
        }
    }

    // ###########################################################################################
    // With none selected, a component's page lists every component's, each with the way to it - which
    // is the page where the others can be read at all.
    // ###########################################################################################
    [Fact]
    public async Task With_no_component_selected_a_page_lists_every_components_and_opens_the_one_it_is_about()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ShowEntryAsync(ContributeEntry.ComponentLinks);

            Assert.True(tab.IsNoComponentPageShownForTests);
            Assert.True(tab.IsEntryProblemsPanelShownForTests);
            Assert.All(tab.EntryProblemRowsForTests, row => Assert.True(row.IsError));
            Assert.Equal(["Open [U8]", "Open [U9]"], tab.EntryProblemRowsForTests.Select(row => row.OpenLabel));
            Assert.All(tab.EntryProblemRowsForTests, row => Assert.True(row.CanOpen));
            Assert.Equal("2 errors here", Title(tab));
            Assert.StartsWith("Of all the board's components.", Help(tab), StringComparison.Ordinal);

            // Open: that component, on this very page.
            tab.EntryProblemRowsForTests[1].OpenCommand.Execute(null);

            for (int attempt = 0; attempt < 200 && !tab.IsComponentOpen; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.True(tab.IsComponentOpen);
            Assert.Equal("U9", tab.ComponentPaneForTests.Label);
            Assert.Equal(ComponentPage.Links, tab.ComponentPaneForTests.Page);
            Assert.Equal(ContributeEntry.ComponentLinks, tab.EntryOpenForTests);
            Assert.Equal("U9", tab.PickedComponentForTests);

            // Its own is what the page has now, with no way to where it already is.
            EntryProblemRow own = Assert.Single(tab.EntryProblemRowsForTests);
            Assert.Equal("U9", own.Component);
            Assert.False(own.CanOpen);
            Assert.Equal("1 error here", Title(tab));
            Assert.DoesNotContain("Of all the board's components", Help(tab), StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // WITH A COMPONENT SELECTED NOTHING OF ANOTHER SHOWS (owner request, 2026-11-04): on its page and on the
    // entries of its pages. "Overview" still has them all, and a board's own entries are the board's.
    // ###########################################################################################
    [Fact]
    public async Task With_a_component_selected_its_entries_and_pages_say_nothing_of_the_others()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            // U9: no category, and a bad link.
            await tab.OpenComponentAsync("U9");
            await tab.ShowEntryAsync(ContributeEntry.ComponentLinks);

            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentBasics));
            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentLinks));
            Assert.Equal(["U9"], tab.EntryProblemRowsForTests.Select(row => row.Component));

            // U8: a bad link only.
            await tab.OpenComponentAsync("U8");

            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentBasics));
            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentLinks));
            Assert.Equal(["U8"], tab.EntryProblemRowsForTests.Select(row => row.Component));

            // C1: nothing wrong, so nothing at all stands over its page - whatever the others have.
            await tab.OpenComponentAsync("C1");

            Assert.Equal(ContributeEntry.ComponentLinks, tab.EntryOpenForTests);
            Assert.False(tab.IsEntryProblemsPanelShownForTests);
            Assert.Empty(tab.EntryProblemRowsForTests);
            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentBasics));
            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentLinks));

            // What is about the board, and the total, do not depend on which component is selected.
            Assert.Equal((string.Empty, "2"), tab.NavBadgesForTests(ContributeEntry.Links));
            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.Credits));
            Assert.Equal(("5", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));

            // Let go of, every component's again.
            tab.LetGoOfPickedComponent();

            Assert.Equal(("2", string.Empty), tab.NavBadgesForTests(ContributeEntry.ComponentLinks));
            Assert.Equal(2, tab.EntryProblemRowsForTests.Count);
        });
    }

    // ###########################################################################################
    // ON EACH COMPONENT IN THE LIST ON THE FAR LEFT (owner request, 2026-11-04): all its pages together, on
    // the row as it stands - made again, the list let go of its selection and scrolled back to it.
    // ###########################################################################################
    [Fact]
    public async Task Each_component_in_the_list_says_its_own_and_follows_a_fix_on_the_row_as_it_stands()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);
            var list = new ListBox { SelectionMode = SelectionMode.Single, ItemsSource = tab.ComponentItems };
            tab.UseComponentList(list);

            Assert.Equal(new BoardProblemCounts(1, 0), Row(tab, "U8").Problems);
            Assert.Equal(new BoardProblemCounts(2, 0), Row(tab, "U9").Problems);
            Assert.False(Row(tab, "C1").HasErrors || Row(tab, "C1").HasWarnings);

            await tab.OpenComponentAsync("U8");
            ComponentListItem shown = Row(tab, "U8");

            Assert.True(shown.HasErrors);
            Assert.Equal("1", shown.ErrorsText);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;

            session.Apply("Links", content => content with
            {
                Board = content.Board.WithComponentSections(componentLinks:
                [
                    new ComponentLinkEntry { BoardLabel = "U8", Name = "Good now", Url = "https://example.com/u8" },
                    content.Board.ComponentLinks.Single(link => link.BoardLabel == "U9"),
                ]),
            });

            Assert.Equal(DraftWorkbookEditOutcome.Saved, await session.SaveAsync());
            File.SetLastWriteTimeUtc(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA), DateTime.UtcNow.AddMinutes(1));

            tab.RefreshDrafts();

            Assert.Same(shown, Row(tab, "U8"));
            Assert.False(shown.HasErrors);
            Assert.Equal(new BoardProblemCounts(2, 0), Row(tab, "U9").Problems);
        });
    }

    [Fact]
    public async Task A_part_of_the_board_says_its_own_and_a_page_with_nothing_wrong_says_nothing()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ShowEntryAsync(ContributeEntry.Links);

            Assert.True(tab.IsEntryProblemsPanelShownForTests);
            Assert.Equal("2 warnings here", Title(tab));
            Assert.Contains("todo", tab.GetControl<Border>("EntryProblemsPanel").Classes);
            Assert.DoesNotContain("error", tab.GetControl<Border>("EntryProblemsPanel").Classes);
            Assert.All(tab.EntryProblemRowsForTests, row => Assert.False(row.CanOpen));

            await tab.ShowEntryAsync(ContributeEntry.Credits);

            Assert.Equal("1 error here", Title(tab));
            Assert.Contains("error", tab.GetControl<Border>("EntryProblemsPanel").Classes);
            Assert.DoesNotContain("todo", tab.GetControl<Border>("EntryProblemsPanel").Classes);
            Assert.StartsWith("An error stops", Help(tab), StringComparison.Ordinal);

            await tab.ShowEntryAsync(ContributeEntry.Files);

            Assert.False(tab.IsEntryProblemsPanelShownForTests);
            Assert.Empty(tab.EntryProblemRowsForTests);
        });
    }

    // A table marks every problem on its cell, with pills of its own: the badges stay, the panel does not.
    [Fact]
    public async Task Over_a_table_the_panel_gives_way_and_the_badges_stay()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ShowEntryAsync(ContributeEntry.Credits);
            Assert.True(tab.IsEntryProblemsPanelShownForTests);

            await tab.ShowViewAsync(table: true);

            Assert.True(tab.IsTableOpen);
            Assert.False(tab.IsEntryProblemsPanelShownForTests);
            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.Credits));

            await tab.ShowViewAsync(table: false);

            Assert.True(tab.IsEntryProblemsPanelShownForTests);
        });
    }

    // Fixed, it goes - from the badge and the page - as the draft's save reads the list again.
    [Fact]
    public async Task What_is_fixed_goes_from_the_badges_and_the_page()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ShowEntryAsync(ContributeEntry.Credits);
            Assert.Equal(("1", string.Empty), tab.NavBadgesForTests(ContributeEntry.Credits));

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;

            session.Apply("Credits", content => content with
            {
                Board = content.Board.WithBoardSections(credits: [new CreditEntry { Category = "Schematic image", SubCategory = "Top", NameOrHandle = "Anna" }]),
            });

            Assert.Equal(DraftWorkbookEditOutcome.Saved, await session.SaveAsync());
            File.SetLastWriteTimeUtc(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA), DateTime.UtcNow.AddMinutes(1));

            tab.RefreshDrafts();

            Assert.Equal((string.Empty, string.Empty), tab.NavBadgesForTests(ContributeEntry.Credits));
            Assert.Equal(("4", "2"), tab.NavBadgesForTests(ContributeEntry.Overview));
            Assert.False(tab.IsEntryProblemsPanelShownForTests);
        });
    }

    // ###########################################################################################
    // In a window: between the entry's name and its count, so the counts stay under each other - an entry
    // with a name alone given a row for them - and the header holding neither text nor badge.
    // ###########################################################################################
    [Fact]
    public void A_badge_stands_between_the_name_and_the_count_and_the_header_has_none()
    {
        BoardData board = Board();
        WriteDraft(board);

        UiTest.Run(() =>
        {
            TabDrafts tab = Tab(board);
            var window = new Window { Content = tab, Width = 1300, Height = 800 };

            window.Show();
            Dispatcher.UIThread.RunJobs();

            Button links = tab.GetControl<Button>("LinksNavButton");
            Grid row = Assert.IsType<Grid>(links.Content);

            TextBlock name = row.Children.OfType<TextBlock>().First(text => text.Text == "Links");
            TextBlock count = tab.GetControl<TextBlock>("LinksCountText");
            Border badge = links.GetVisualDescendants().OfType<Border>().First(border => border.Classes.Contains("NavBadge") && border.IsVisible);

            double X(Control control) => control.TranslatePoint(default, links)!.Value.X;

            Assert.True(X(name) + name.Bounds.Width <= X(badge), "The badge should stand after the name");
            Assert.True(X(badge) + badge.Bounds.Width <= X(count), "The badge should stand before the count");
            Assert.Contains("warning", badge.Classes);

            // "Overview" and "Name" have no count of their own: a row is made for them, and still reads so.
            Button overview = tab.GetControl<Button>("StatusNavButton");
            Assert.IsType<Grid>(overview.Content);
            Assert.Contains(overview.GetVisualDescendants().OfType<TextBlock>(), text => text.Text == "Overview");
            Assert.Equal(2, overview.GetVisualDescendants().OfType<Border>().Count(border => border.Classes.Contains("NavBadge") && border.IsVisible));

            // Nothing in the header says them any more.
            List<string> header = tab.GetControl<ContentControl>("CurrentDraftRow").GetVisualDescendants()
                .OfType<TextBlock>()
                .Where(text => text.IsEffectivelyVisible)
                .Select(text => text.Text ?? string.Empty)
                .ToList();

            Assert.DoesNotContain(header, text => text.Contains("error", StringComparison.OrdinalIgnoreCase) || text.Contains("warning", StringComparison.OrdinalIgnoreCase));

            window.Close();
        });
    }

    // What the checks found holds the draft's whole board: a draft gone is let go of by the next refresh of
    // the list (code review, 2026-11-04), which runs after every discard and retirement.
    [Fact]
    public async Task What_the_checks_found_of_a_draft_gone_is_let_go_of_by_the_next_refresh()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);
            string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA);

            Assert.True(DraftStatusReader.HoldsProblemsForTests(boardFile));

            // Windows refuses to delete a file the tab's background read still holds.
            await tab.ReadinessShownForTests;

            // The whole folder, as a discard takes it: with its marker left it is still a draft to read.
            Directory.Delete(DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, BoardA), recursive: true);
            tab.RefreshDrafts();

            Assert.False(DraftStatusReader.HoldsProblemsForTests(boardFile));
        });
    }

    // ###########################################################################################
    // On the way back to "Overview" the last checks are shown at once, until the new ones are in - and their
    // row numbers are of the board they ran on (code review, 2026-11-04: read on the newer board, U9's
    // problem was said of the wrong row, and the first could open what was not meant). So each stays about
    // its own row, a component gone has no way to it, and the new checks then say what stands.
    // ###########################################################################################
    [Fact]
    public async Task What_Overview_says_before_the_new_checks_are_in_is_of_the_board_they_ran_on()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ReadinessShownForTests;

            // Each bad address and the way beside it: "u8: Open [U8]".
            List<string> Lines() => tab.StatusProblemsForTests(BoardStatusPart.Components)!.Value.Rows
                .Where(row => row.Children.OfType<TextBlock>().Any(text => text.Text?.Contains("ftp://", StringComparison.Ordinal) == true))
                .Select(row =>
                {
                    string about = row.Children.OfType<TextBlock>().Any(text => text.Text?.Contains("/u8", StringComparison.Ordinal) == true) ? "u8" : "u9";
                    string way = row.Children.OfType<Button>().SingleOrDefault()?.Content as string ?? "no way";

                    return $"{about}: {way}";
                })
                .ToList();

            Assert.Equal(["u8: Open [U8]", "u9: Open [U9]"], Lines());

            // U8 goes, with its link: U9's link is the first row now.
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;

            Assert.True(session.Apply("U8 deleted", content => content with
            {
                Board = content.Board.WithComponentSections(
                    components: content.Board.Components.Where(row => row.BoardLabel != "U8").ToList(),
                    componentLinks: content.Board.ComponentLinks.Where(row => row.BoardLabel != "U8").ToList()),
            }));

            // What shows at once, nothing having been waited for.
            tab.ShowBoard(Entry(), session.Content.Board, "PAL", null);

            Assert.Equal(["u8: no way", "u9: Open [U9]"], Lines());

            // The new checks in: what stands.
            await tab.ReadinessShownForTests;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["u9: Open [U9]"], Lines());
        });
    }

    // ###########################################################################################
    // EACH PANEL OF "OVERVIEW" HAS ITS OWN IN A PANEL OF THEIR OWN, RIGHT UNDER ITS HEADLINE (owner requests,
    // 2026-11-04: "the error and warning panels should be shown specifically inside those panels and then
    // detail the issues in there also"; "encapsulated in its own panel, as now the text is kind of mixed with the
    // normal text"): how many, every line, and from a component's the way to it.
    // ###########################################################################################
    [Fact]
    public async Task Each_panel_of_Overview_has_its_errors_and_warnings_in_a_panel_of_their_own()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ReadinessShownForTests;

            var credits = tab.StatusProblemsForTests(BoardStatusPart.Credits)!.Value;

            Assert.Equal("1 error", credits.Title);
            Assert.Contains("error", tab.StatusProblemsBlockForTests(BoardStatusPart.Credits).Classes);

            Grid credit = Assert.Single(credits.Rows);
            List<string> creditTexts = credit.Children.OfType<TextBlock>().Select(text => text.Text ?? string.Empty).ToList();

            // Its row named, and its field in quotes as a field is.
            Assert.Contains(creditTexts, text => text.StartsWith("[Schematic image / Top] This credit has no \"Name or handle\".", StringComparison.Ordinal));
            Assert.Empty(credit.Children.OfType<Button>());

            // Right under the headline and what it explains, above what the panel counts.
            var inPanel = (StackPanel)tab.StatusProblemsBlockForTests(BoardStatusPart.Credits).Parent!;

            Assert.Equal(2, inPanel.Children.IndexOf(tab.StatusProblemsBlockForTests(BoardStatusPart.Credits)));

            // "Show in the table" is about what they say, so it stands with them.
            Assert.Same(
                tab.StatusProblemsBlockForTests(BoardStatusPart.Credits),
                tab.StatusTableButtonForTests(BoardStatusPart.Credits).GetLogicalAncestors().OfType<Border>().First());

            var links = tab.StatusProblemsForTests(BoardStatusPart.Links)!.Value;

            Assert.Equal("2 warnings", links.Title);
            Assert.Contains("todo", tab.StatusProblemsBlockForTests(BoardStatusPart.Links).Classes);
            Assert.Equal(2, links.Rows.Count);

            var components = tab.StatusProblemsForTests(BoardStatusPart.Components)!.Value;

            Assert.Equal("2 errors", components.Title);

            List<string> ways = components.Rows.SelectMany(row => row.Children.OfType<Button>()).Select(button => button.Content as string ?? string.Empty).ToList();

            Assert.Equal(["Open [U8]", "Open [U9]"], ways);

            // Each of those has the way to its component, so no table (owner request, 2026-11-04).
            Assert.False(tab.StatusTableButtonForTests(BoardStatusPart.Components).IsVisible);
            Assert.True(tab.StatusTableButtonForTests(BoardStatusPart.Credits).IsVisible);

            // Nothing wrong, no panel.
            Assert.Null(tab.StatusProblemsForTests(BoardStatusPart.Name));

            // The way to the component is that component, on the page it is on.
            components.Rows[1].Children.OfType<Button>().Single().RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            for (int attempt = 0; attempt < 200 && !tab.IsComponentOpen; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.True(tab.IsComponentOpen, "The way to the component should open it");
            Assert.Equal("U9", tab.ComponentPaneForTests.Label);
            Assert.Equal(ComponentPage.Links, tab.ComponentPaneForTests.Page);
        });
    }

    // The components marked nowhere: in that panel, counted, one line and a button each.
    [Fact]
    public async Task The_components_marked_nowhere_stand_in_the_panel_of_the_warnings()
    {
        BoardData board = Board();
        board.ComponentHighlights.RemoveAll(row => row.BoardLabel == "C1");
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ReadinessShownForTests;

            var components = tab.StatusProblemsForTests(BoardStatusPart.Components)!.Value;

            Assert.Equal("2 errors and 1 warning", components.Title);
            Assert.DoesNotContain(
                components.Rows.SelectMany(row => row.Children.OfType<TextBlock>()),
                text => (text.Text ?? string.Empty).Contains("[C1]", StringComparison.Ordinal));

            // Its other groups - no short description, say - stand under it, as they are no warnings.
            List<Button> ofC1 = tab.GetControl<StackPanel>("StatusPanelsHost")
                .GetLogicalDescendants()
                .OfType<Button>()
                .Where(button => (button.Tag as string) == "C1")
                .ToList();

            Assert.Contains(ofC1, button => button.GetLogicalAncestors().OfType<Border>().Any(border => border.Classes.Contains("StatusProblems")));

            // A sentence in that panel, under its title - not a heading of its own (owner request, 2026-11-04).
            TextBlock sentence = tab.GetControl<StackPanel>("StatusPanelsHost")
                .GetLogicalDescendants()
                .OfType<TextBlock>()
                .Single(text => text.Text == BoardReadiness.UnmarkedHeading(1));

            Assert.StartsWith("The following component is not marked on any schematic", sentence.Text, StringComparison.Ordinal);
            Assert.Equal(Avalonia.Media.FontWeight.Normal, sentence.FontWeight);
        });
    }

    // ###########################################################################################
    // ONE PHYSICAL COMPONENT, ONE WARNING, WHEREVER IT IS COUNTED (owner decision, 2026-11-04: "it should only
    // be reported once as it is ONE physical component only"). A row per region read as two: the page said the
    // same sentence twice, and the panel's title disagreed with the sentence under it.
    // ###########################################################################################
    [Fact]
    public async Task A_component_with_a_row_per_region_marked_nowhere_counts_once_everywhere()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
            Components =
            [
                new ComponentEntry { BoardLabel = "U19", Region = "PAL", FriendlyName = "VIC-II", Category = "IC" },
                new ComponentEntry { BoardLabel = "U19", Region = "NTSC", FriendlyName = "VIC-II", Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", FriendlyName = "Filter", Category = "Capacitor" },
            ],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "C1", X = "1", Y = "1", Width = "5", Height = "5" }],
        };

        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ReadinessShownForTests;

            Assert.Equal((string.Empty, "1"), tab.NavBadgesForTests(ContributeEntry.ComponentBasics));
            Assert.Equal(new BoardProblemCounts(0, 1), Row(tab, "U19").Problems);
            Assert.Equal("1 warning", tab.StatusProblemsForTests(BoardStatusPart.Components)!.Value.Title);
            Assert.Contains(
                tab.GetControl<StackPanel>("StatusPanelsHost").GetLogicalDescendants().OfType<TextBlock>(),
                text => text.Text == BoardReadiness.UnmarkedHeading(1));

            await tab.OpenComponentAsync("U19");
            await tab.ShowEntryAsync(ContributeEntry.ComponentBasics);

            Assert.Equal("1 warning here", Title(tab));
            Assert.Single(tab.EntryProblemRowsForTests);
        });
    }

    // ###########################################################################################
    // A PAGE'S PANEL STANDS UNDER ITS HEADLINE (owner request, 2026-11-04: "should come below the headline
    // everywhere"), and under a component's buttons, so those do not move as a component with issues is
    // followed by one without. One panel, moved to the page open.
    // ###########################################################################################
    [Fact]
    public async Task A_pages_panel_stands_under_its_headline_and_a_components_buttons_stay_where_they_are()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);
            var window = new Window { Content = tab, Width = 1300, Height = 800 };

            window.Show();
            Dispatcher.UIThread.RunJobs();

            Border panel = tab.GetControl<Border>("EntryProblemsPanel");
            Decorator noComponent = tab.GetControl<Decorator>("NoComponentProblemsHost");
            ComponentPane pane = tab.ComponentPaneForTests;

            double Y(Control control) => control.TranslatePoint(default, tab)!.Value.Y;

            // A part of the board, in the pane.
            await tab.ShowEntryAsync(ContributeEntry.Credits);
            Dispatcher.UIThread.RunJobs();

            Assert.True(panel.IsEffectivelyVisible);
            Assert.Same(pane.ProblemsHost, tab.EntryProblemsHostForTests);
            Assert.True(Y(panel) > Y(pane.GetControl<TextBlock>("TitleText")), "The panel should stand under the headline");

            // A component's page with none picked.
            tab.LetGoOfPickedComponent();
            await tab.ShowEntryAsync(ContributeEntry.ComponentLinks);
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.IsNoComponentPageShownForTests);
            Assert.True(panel.IsEffectivelyVisible);
            Assert.Same(noComponent, tab.EntryProblemsHostForTests);
            Assert.True(Y(panel) > Y(tab.GetControl<TextBlock>("NoComponentTitle")), "The panel should stand under the headline");

            // A component with issues, and one without: its buttons stand in the same place.
            await tab.OpenComponentAsync("U8");
            Dispatcher.UIThread.RunJobs();

            Assert.True(panel.IsEffectivelyVisible);
            Assert.Same(pane.ProblemsHost, tab.EntryProblemsHostForTests);

            // A button of the header's - "Delete component" stands on "Basics" since 2026-11-07, under the panel.
            Control button = pane.GetControl<Button>("PreviewPopupButton");
            double withIssues = Y(button);

            Assert.True(Y(panel) > withIssues, "The panel should stand under the buttons");

            await tab.OpenComponentAsync("C1");
            Dispatcher.UIThread.RunJobs();

            Assert.False(panel.IsVisible);
            Assert.Equal(withIssues, Y(button));

            window.Close();
        });
    }

    // ---- "With issues" ----------------------------------------------------------

    private static (ListBox List, CheckBox Box) ListAndBoxFor(TabDrafts tab)
    {
        var list = new ListBox { SelectionMode = SelectionMode.Single, ItemsSource = tab.ComponentItems };
        var box = new CheckBox();

        tab.UseIssuesBox(box);
        tab.UseComponentList(list);

        return (list, box);
    }

    private static List<string> Listed(TabDrafts tab) => tab.ComponentItems.Select(item => item.BoardLabel).ToList();

    // A save of `change`, as far as the list hears of it.
    private static async Task SaveAndReadAsync(TabDrafts tab, Func<BoardContent, BoardContent?> change)
    {
        DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardA)!;

        session.Apply("A change", change);

        Assert.Equal(DraftWorkbookEditOutcome.Saved, await session.SaveAsync());
        File.SetLastWriteTimeUtc(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA), DateTime.UtcNow.AddMinutes(1));

        tab.RefreshDrafts();
    }

    private static BoardContent WithLinks(BoardContent content, params ComponentLinkEntry[] links) =>
        content with { Board = content.Board.WithComponentSections(componentLinks: [.. links]) };

    private static ComponentLinkEntry Good(string label) => new() { BoardLabel = label, Name = "Good now", Url = "https://example.com/" + label };

    private static ComponentLinkEntry Bad(string label) => new() { BoardLabel = label, Name = "Bad", Url = "ftp://example.com/" + label };

    // ###########################################################################################
    // ONLY THOSE WITH AN ERROR OR A WARNING, EACH BOARD ITS OWN (owner requests, 2026-11-04): a board with
    // nothing wrong has no box at all, and the tick waits for its own board.
    // ###########################################################################################
    [Fact]
    public void The_box_narrows_the_list_to_those_with_issues_for_the_board_it_was_ticked_on()
    {
        BoardData board = Board();
        WriteDraft(board);

        UiTest.Run(() =>
        {
            TabDrafts tab = Tab(board);
            (ListBox _, CheckBox box) = ListAndBoxFor(tab);

            Assert.True(box.IsVisible);
            Assert.False(box.IsChecked);
            Assert.Equal(["U8", "U9", "C1"], Listed(tab));

            box.IsChecked = true;

            Assert.Equal(["U8", "U9"], Listed(tab));

            tab.ShowBoard(Entry("250425", BoardB), new BoardData { Components = [new ComponentEntry { BoardLabel = "R1", Category = "Resistor" }] }, "PAL", null);

            Assert.Equal(["R1"], Listed(tab));
            Assert.False(box.IsVisible);
            Assert.False(box.IsChecked);

            tab.ShowBoard(Entry(), board, "PAL", null);

            Assert.True(box.IsVisible);
            Assert.True(box.IsChecked);
            Assert.Equal(["U8", "U9"], Listed(tab));

            box.IsChecked = false;

            Assert.Equal(["U8", "U9", "C1"], Listed(tab));
        });
    }

    // ###########################################################################################
    // THE LIST FILLS IN ONE GO AND COMES BACK TO THE COMPONENT PICKED ONCE (owner report, 2026-11-04: "some kind
    // of slow scroll behaviour, where it populates the list and scrolls to the same component afterwards"). Every
    // row put in above the selection moves its index, and the list asked to be scrolled to each index in turn, a
    // layout pass apiece - which is seen as the list scrolling by itself. A shown window: nothing scrolls in
    // one that was never laid out.
    // ###########################################################################################
    [Fact]
    public async Task Unticking_the_box_fills_the_list_in_one_go_and_comes_back_to_the_component_once()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
            Components = [.. Enumerable.Range(1, 400).Select(number => new ComponentEntry { BoardLabel = $"U{number}", FriendlyName = "Chip", Category = "IC" })],
            ComponentHighlights = [.. Enumerable.Range(1, 400).Select(number => new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = $"U{number}", X = "1", Y = "1", Width = "5", Height = "5" })],
            ComponentLinks = [Bad("U400")],
        };

        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);
            (ListBox list, CheckBox box) = ListAndBoxFor(tab);
            var window = new Window { Content = list, Width = 300, Height = 300 };

            window.Show();
            Dispatcher.UIThread.RunJobs();

            box.IsChecked = true;
            tab.PickComponentForTests("U400");

            for (int attempt = 0; attempt < 200 && !tab.IsComponentOpen; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["U400"], Listed(tab));

            ScrollViewer scroller = list.GetVisualDescendants().OfType<ScrollViewer>().First();
            int moves = 0;

            scroller.PropertyChanged += (_, change) =>
            {
                if (change.Property == ScrollViewer.OffsetProperty)
                {
                    moves++;
                }
            };

            box.IsChecked = false;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(400, tab.ComponentItems.Count);
            Assert.Equal("U400", (list.SelectedItem as ComponentListItem)?.BoardLabel);

            // Brought into view, and not walked to.
            Assert.NotNull(list.ContainerFromIndex(399));
            Assert.InRange(moves, 1, 2);

            window.Close();
        });
    }

    // Picked with nothing wrong as the box is ticked it goes from the list, and the pane stays on it - as for
    // a category unticked.
    [Fact]
    public async Task A_component_with_nothing_wrong_goes_from_the_list_as_the_box_is_ticked()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);
            (ListBox list, CheckBox box) = ListAndBoxFor(tab);

            await tab.OpenComponentAsync("C1");
            Assert.Equal("C1", (list.SelectedItem as ComponentListItem)?.BoardLabel);

            box.IsChecked = true;

            Assert.Equal(["U8", "U9"], Listed(tab));
            Assert.Null(list.SelectedItem);
            Assert.Equal("C1", tab.ComponentPaneForTests.Label);
            Assert.Equal("C1", tab.PickedComponentForTests);
        });
    }

    // ###########################################################################################
    // WHAT IS BEING FIXED STAYS LISTED UNTIL ANOTHER IS PICKED. Its row going a second after the fix would take
    // the selection with it, under the one working on it.
    // ###########################################################################################
    [Fact]
    public async Task A_component_fixed_stays_listed_until_another_is_picked()
    {
        BoardData board = Board();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);
            (ListBox list, CheckBox box) = ListAndBoxFor(tab);

            box.IsChecked = true;
            tab.PickComponentForTests("U8");

            for (int attempt = 0; attempt < 200 && !tab.IsComponentOpen; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            await SaveAndReadAsync(tab, content => WithLinks(content, Good("U8"), Bad("U9")));

            Assert.False(Row(tab, "U8").HasErrors);
            Assert.Equal(["U8", "U9"], Listed(tab));
            Assert.Equal("U8", (list.SelectedItem as ComponentListItem)?.BoardLabel);

            tab.PickComponentForTests("U9");

            for (int attempt = 0; attempt < 200 && tab.ComponentPaneForTests.Label != "U9"; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal("U9", tab.ComponentPaneForTests.Label);
            Assert.Equal(["U9"], Listed(tab));
            Assert.Equal("U9", (list.SelectedItem as ComponentListItem)?.BoardLabel);
        });
    }

    // ###########################################################################################
    // THE LAST ISSUE GONE TAKES THE TICK WITH IT. Left ticked underneath, the next warning - any new component
    // has one - would narrow the list by itself.
    // ###########################################################################################
    [Fact]
    public async Task The_last_issue_gone_lets_go_of_the_tick_for_good()
    {
        // Nothing wrong on any component but U8's link.
        BoardData board = Board();
        board = board.WithComponentSections(components: [.. board.Components.Select(row => row.BoardLabel == "U9" ? new ComponentEntry { BoardLabel = "U9", FriendlyName = "VIC", Category = "IC" } : row)]);
        board = board.WithComponentSections(componentLinks: [Bad("U8")]);
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);
            (ListBox _, CheckBox box) = ListAndBoxFor(tab);

            box.IsChecked = true;
            Assert.Equal(["U8"], Listed(tab));

            await SaveAndReadAsync(tab, content => WithLinks(content, Good("U8")));

            Assert.False(box.IsVisible);
            Assert.False(box.IsChecked);
            Assert.Equal(["U8", "U9", "C1"], Listed(tab));

            await SaveAndReadAsync(tab, content => WithLinks(content, Bad("C1")));

            Assert.True(box.IsVisible);
            Assert.False(box.IsChecked);
            Assert.Equal(["U8", "U9", "C1"], Listed(tab));
        });
    }

    // ---- "Open" ------------------------------------------------------------------

    // Waits for what "Open" does, which runs under the tab's own guard.
    private static async Task UntilAsync(Func<bool> done)
    {
        for (int attempt = 0; attempt < 200 && !done(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(done(), "What was opened never came");
    }

    // A chip with a row per region, and a baseline of the NTSC one off its case.
    private static BoardData BoardWithAnNtscBaselineOffItsCase() => new()
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

    // ###########################################################################################
    // "OPEN [U7]" ON A PROBLEM OF THE REGION NOT SHOWN SHOWS THAT REGION (code review, 2026-11-04). A page
    // lists the region shown, so with CRT on PAL the NTSC baseline the badges counted was on no page: no row
    // looked wrong, the entry was not red, and the owner's "it is not at all clear for me where those errors
    // are" held. And what is said of it names the region, though its message names its row already.
    // ###########################################################################################
    [Fact]
    public async Task Open_on_a_problem_of_the_region_not_shown_shows_the_row_it_is_about()
    {
        BoardData board = BoardWithAnNtscBaselineOffItsCase();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ShowEntryAsync(ContributeEntry.ComponentBaselines);

            EntryProblemRow row = tab.EntryProblemRowsForTests.Single(item => item.Message.Contains("pin 99", StringComparison.Ordinal));

            Assert.Contains("NTSC", row.Message, StringComparison.Ordinal);
            Assert.True(row.CanOpen);

            row.OpenCommand.Execute(null);

            await UntilAsync(() => tab.IsComponentOpen && tab.ComponentPaneForTests.BaselinesForTests.Count > 0);

            Assert.Equal("U7", tab.ComponentPaneForTests.Label);
            Assert.Equal(ComponentPage.Baselines, tab.ComponentPaneForTests.Page);
            Assert.Equal("99", Assert.Single(tab.ComponentPaneForTests.BaselinesForTests).Pin);
        });
    }

    // On the component already, on the wrong region: the way to the row is still needed, and there it stops.
    [Fact]
    public async Task With_the_pane_on_the_component_in_the_other_region_Open_still_goes_to_the_row()
    {
        BoardData board = BoardWithAnNtscBaselineOffItsCase();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.OpenComponentAsync("U7");
            await tab.ShowEntryAsync(ContributeEntry.ComponentBaselines);

            Assert.Empty(tab.ComponentPaneForTests.BaselinesForTests);

            EntryProblemRow row = tab.EntryProblemRowsForTests.Single(item => item.Message.Contains("pin 99", StringComparison.Ordinal));

            Assert.True(row.CanOpen);

            row.OpenCommand.Execute(null);

            await UntilAsync(() => tab.ComponentPaneForTests.BaselinesForTests.Count > 0);

            Assert.False(tab.EntryProblemRowsForTests.Single(item => item.Message.Contains("pin 99", StringComparison.Ordinal)).CanOpen);
        });
    }

    // ###########################################################################################
    // A ROW KEPT IN PLACE OPENS WHAT IT NOW SAYS (code review, 2026-11-04). The same message on two pages - a
    // picture and a baseline of one component citing one file - was kept as "the same" going from one page
    // to the other, its "Open" still going to the page left behind.
    // ###########################################################################################
    [Fact]
    public async Task A_row_saying_the_same_on_another_page_opens_that_page()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
            Components = [new ComponentEntry { BoardLabel = "U7", FriendlyName = "VIC", Category = "IC" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U7", X = "1", Y = "1", Width = "5", Height = "5" }],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U7", Name = "Photo", File = "Images/both.pdf" },
                new ComponentImageEntry { BoardLabel = "U7", Pin = "1", Name = "Clock", File = "Images/both.pdf" },
            ],
        };

        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(board);

            await tab.ShowEntryAsync(ContributeEntry.ComponentPictures);
            Assert.NotEmpty(tab.EntryProblemRowsForTests);

            await tab.ShowEntryAsync(ContributeEntry.ComponentBaselines);
            Assert.NotEmpty(tab.EntryProblemRowsForTests);

            tab.EntryProblemRowsForTests[0].OpenCommand.Execute(null);

            await UntilAsync(() => tab.IsComponentOpen);

            Assert.Equal(ComponentPage.Baselines, tab.ComponentPaneForTests.Page);
        });
    }
}
