using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A new board's credits start with their rows, where the contributor goes to them - "Board: Credits",
// the table and Submit - and "Board status" says how many are to come (owner request, 2026-11-02; the rule
// is NewBoardCreditsTests'). A board with a published counterpart gets none, and a row deleted stays
// deleted, also for the next run of CRT.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsNewBoardCreditsTests : IDisposable
{
    private const string Identity = "Commodore/C64/999001/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private static readonly HardwareBoardEntry Entry = new()
    {
        HardwareName = "Commodore 64",
        BoardName = "999001",
        ExcelDataFile = TabDraftsNewBoardCreditsTests.Identity,
        IsDraftOnly = true,
    };

    public TabDraftsNewBoardCreditsTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static void SeedDraft(bool newBoard)
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);

        CachedWorkbooks.Write(boardFile, new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" }],
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
        });

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity),
            new DraftMarker
            {
                BoardKey = TabDraftsNewBoardCreditsTests.Identity,
                NewBoard = newBoard
                    ? new NewBoardRegistration { ExcelDataFile = TabDraftsNewBoardCreditsTests.Identity, HardwareName = "Commodore 64", BoardName = "999001" }
                    : null,
            });
    }

    private static TabDrafts Tab()
    {
        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [TabDraftsNewBoardCreditsTests.Entry],
            PublishedBoardOverrideForTests = _ => null,
        };

        tab.RefreshDrafts();
        tab.ShowBoard(TabDraftsNewBoardCreditsTests.Entry, new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] }, "PAL", null);

        return tab;
    }

    private static List<string> Credits() =>
        DraftManager.Sessions.Get(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity)!.Content.Board.Credits
            .Select(row => $"{row.Category}|{row.SubCategory}")
            .ToList();

    private static readonly string[] Wanted =
        ["Schematic image|Top", "Board data|", "Oscilloscope baseline|PAL", "Oscilloscope baseline|NTSC"];

    [Fact]
    public async Task Opening_the_credits_of_a_new_board_shows_its_rows_as_one_step_and_undo_takes_them()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(TabDraftsNewBoardCreditsTests.Wanted, tab.ComponentPaneForTests.CreditsForTests.Select(row => $"{row.Category}|{row.SubCategory}"));
            Assert.All(tab.ComponentPaneForTests.CreditsForTests, row => Assert.True(row.NameOrHandleMissing));

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity)!;
            Assert.Equal(DraftStepWording.CreditsOfNewBoard, session.UndoDescription);

            // Undone, they stay gone: what was offered is in the marker.
            session.Undo();
            await tab.OpenBoardPartAsync(PaneBoardPart.Links);
            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(tab.ComponentPaneForTests.CreditsForTests);

            tab.CloseComponent();
        });

        Assert.Equal(
            TabDraftsNewBoardCreditsTests.Wanted.Select(row => row.ToUpperInvariant()).Order(StringComparer.Ordinal),
            DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity))!.CreditsOffered!);
    }

    // ###########################################################################################
    // *** READING "BOARD STATUS" CHANGES NOTHING IN THE DRAFT (code review, 2026-11-02). *** Worked out
    // again after every change, so a read that wrote was behind the Undo below and the rename race: it
    // says how many rows are to come, hands nothing over, adds no step and writes no marker.
    // ###########################################################################################
    [Fact]
    public async Task Board_status_leaves_the_draft_as_it_is_and_says_what_opening_the_credits_adds()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();
            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity)!;

            int handedOver = 0;
            session.ChangeComing += (_, _) => handedOver++;

            await tab.ShowBoardStatusAsync();
            await tab.ReadinessShownForTests;
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(Credits());
            Assert.False(session.CanUndo);
            Assert.Equal(0, handedOver);
            Assert.Null(DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity))!.CreditsOffered);

            BoardStatusPanel panel = tab.StatusPanelForTests(BoardStatusPart.Credits);
            Assert.Equal(BoardStatusState.ToDo, panel.State);
            Assert.Contains(panel.Lines, line => line.Text == BoardReadiness.CreditsToFillIn(4));

            // Gone to, they are there, and the panel no longer says they are to come.
            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(TabDraftsNewBoardCreditsTests.Wanted, Credits());

            await tab.ShowBoardStatusAsync();
            await tab.ReadinessShownForTests;
            Dispatcher.UIThread.RunJobs();

            panel = tab.StatusPanelForTests(BoardStatusPart.Credits);
            Assert.Equal(BoardStatusState.Error, panel.State);
            Assert.DoesNotContain(panel.Lines, line => line.Text == BoardReadiness.CreditsToFillIn(4));
        });
    }

    // ###########################################################################################
    // A read that outlives its page still touches nothing (found by the Stop hook's run, 2026-11-02:
    // "A_new_boards_name_is_changed_on_its_page" failed at random, the marker written while "Board name"
    // moved the draft's folder).
    // ###########################################################################################
    [Fact]
    public async Task A_status_read_that_outlives_its_page_leaves_the_draft_alone()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            // Nothing below yields before the page has gone, so the read cannot come back any earlier.
            TabDrafts tab = Tab();
            await tab.OpenBoardNameAsync();

            await tab.ReadinessShownForTests;
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(Credits());
            Assert.Null(DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity))!.CreditsOffered);
        });
    }

    // The table shows the Credits sheet, so they are there as it opens - as ONE step, and once.
    [Fact]
    public async Task Opening_the_table_gives_a_new_board_its_rows_too()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();
            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity)!;

            await tab.OpenTableAsync(TabDraftsNewBoardCreditsTests.Entry);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(TabDraftsNewBoardCreditsTests.Wanted, Credits());
            Assert.Equal(DraftStepWording.CreditsOfNewBoard, session.UndoDescription);

            await tab.CloseTableAsync();
            Dispatcher.UIThread.RunJobs();

            // Opened again with nothing new, it adds no step on top of the latest.
            session.Apply("Something else", content => content with { Board = content.Board.WithRegions(["PAL", "NTSC"]) });

            await tab.OpenTableAsync(TabDraftsNewBoardCreditsTests.Entry);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(TabDraftsNewBoardCreditsTests.Wanted, Credits());
            Assert.Equal("Something else", session.UndoDescription);

            await tab.CloseTableAsync();
        });
    }

    // Offered with nothing open beside the board, as before Submit: the left column counts them at once
    // (code review, 2026-11-02 - it went by the board Main last read, until the next read of it).
    [Fact]
    public async Task The_left_column_counts_the_rows_as_they_come_with_nothing_open()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();
            await tab.ReadinessShownForTests;
            Dispatcher.UIThread.RunJobs();

            TextBlock count = tab.GetControl<TextBlock>("CreditsCountText");
            Assert.Equal(ContributeBoardParts.CountText(0), count.Text);

            await tab.PrepareOtherDraftAsync(TabDraftsNewBoardCreditsTests.Entry, string.Empty, string.Empty);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(TabDraftsNewBoardCreditsTests.Wanted, Credits());
            Assert.Equal(ContributeBoardParts.CountText(4), count.Text);
        });
    }

    // Opened again with nothing new: no step, however often, and the Redo stands.
    [Fact]
    public async Task Opening_the_credits_with_nothing_new_is_no_step()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();
            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity)!;

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            await tab.OpenBoardPartAsync(PaneBoardPart.Links);

            session.Apply("Something", content => content with { Board = content.Board.WithRegions(["PAL", "NTSC"]) });
            Assert.True(session.Undo());
            Assert.True(session.CanRedo);

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();

            Assert.True(session.CanRedo);

            tab.CloseComponent();
        });
    }

    // ###########################################################################################
    // *** UNDO ON "BOARD STATUS" WALKS BACK PAST THE CREDITS' STEP (code review, 2026-11-02). *** The
    // undo read the board again, the page was worked out again, and that offered the very step back, the
    // Redo gone with it: an image's rename or delete could never be undone from here, though the delete's
    // own prompt promises "Undo" brings it back.
    // ###########################################################################################
    [Fact]
    public async Task Undo_on_board_status_walks_back_past_the_credits_step()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();
            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity)!;

            // What Main does as a draft changes: the board read again, handed back, the page worked out.
            tab.ReloadBoardOverrideForTests = _ => tab.ShowBoard(TabDraftsNewBoardCreditsTests.Entry, session.Content.Board, "PAL", null);

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            await tab.OpenBoardPartAsync(PaneBoardPart.Links);

            session.Apply("Renamed", content => content with
            {
                Board = content.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "top.png" }]),
            });

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(DraftStepWording.CreditsOfNewBoard, session.UndoDescription);

            await tab.ShowBoardStatusAsync();
            await tab.ReadinessShownForTests;
            Dispatcher.UIThread.RunJobs();

            await tab.UndoAsync(TabDraftsNewBoardCreditsTests.Entry);
            await tab.ReadinessShownForTests;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Renamed", session.UndoDescription);
            Assert.True(session.CanRedo);

            await tab.UndoAsync(TabDraftsNewBoardCreditsTests.Entry);
            await tab.ReadinessShownForTests;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Top", session.Content.Board.Schematics.Single().SchematicName);
            Assert.Contains("Schematic image|Top", Credits());
            Assert.True(session.CanRedo);
        });
    }

    // ###########################################################################################
    // *** ONLY THE STEP UNDONE (code review, 2026-11-02). *** An image renamed, and the credits' step that
    // followed it undone: the rows came back with the old name's row - no longer "offered", so left for
    // good - and the new name's row "offered", so never added. Submit stopped on a credit for an image
    // that no longer had that name. What was offered goes back with such a step, so it follows again.
    // ###########################################################################################
    [Fact]
    public async Task The_step_that_followed_a_renamed_image_undone_follows_it_again()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            await tab.OpenBoardPartAsync(PaneBoardPart.Links);
            Dispatcher.UIThread.RunJobs();

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity)!;

            session.Apply("Renamed", content => content with
            {
                Board = content.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "top.png" }]),
            });

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();

            Assert.Contains("Schematic image|Bottom", Credits());
            Assert.DoesNotContain("Schematic image|Top", Credits());

            // The step alone is undone, then the credits are looked at again.
            Assert.Equal(DraftStepWording.CreditsOfNewBoard, session.UndoDescription);
            session.Undo();
            Assert.Contains("Schematic image|Top", Credits());

            await tab.OpenBoardPartAsync(PaneBoardPart.Links);
            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();

            Assert.Contains("Schematic image|Bottom", Credits());
            Assert.DoesNotContain("Schematic image|Top", Credits());

            tab.CloseComponent();
        });
    }

    // Done again by Redo, what was offered is as that step left it - no row offered twice.
    [Fact]
    public async Task The_step_redone_leaves_what_was_offered_as_it_was_after_it()
    {
        SeedDraft(newBoard: true);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            await tab.OpenBoardPartAsync(PaneBoardPart.Links);
            Dispatcher.UIThread.RunJobs();

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity)!;

            session.Apply("Renamed", content => content with
            {
                Board = content.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "top.png" }]),
            });

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();

            string marker = DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, TabDraftsNewBoardCreditsTests.Identity);
            IReadOnlyList<string> after = DraftMarkerStore.Load(marker)!.CreditsOffered!;

            session.Undo();
            Assert.Contains("SCHEMATIC IMAGE|TOP", DraftMarkerStore.Load(marker)!.CreditsOffered!);

            session.Redo();
            Assert.Equal(after, DraftMarkerStore.Load(marker)!.CreditsOffered!);

            tab.CloseComponent();
        });
    }

    // ###########################################################################################
    // The marker can still say "new board" once the board is published here - approved and synced, its
    // draft not retired yet (code review, 2026-11-02). Board status and Submit ask the entry too; "Board:
    // Credits" did not, and put an unnamed row into that draft, which stopped its Submit.
    // ###########################################################################################
    [Fact]
    public async Task A_new_boards_draft_whose_board_is_published_by_now_gets_none()
    {
        SeedDraft(newBoard: true);

        var published = new HardwareBoardEntry
        {
            HardwareName = "Commodore 64",
            BoardName = "999001",
            ExcelDataFile = TabDraftsNewBoardCreditsTests.Identity,
            IsDraftOnly = false,
        };

        await UiTest.RunAsync(async () =>
        {
            var tab = new TabDrafts
            {
                HardwareBoardsOverrideForTests = [published],
                PublishedBoardOverrideForTests = _ => new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] },
            };

            tab.RefreshDrafts();
            tab.ShowBoard(published, new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] }, "PAL", null);

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(tab.ComponentPaneForTests.CreditsForTests);

            tab.CloseComponent();
        });

        Assert.Empty(Credits());
    }

    // A draft of a published board has its credits as published.
    [Fact]
    public async Task A_draft_of_a_published_board_gets_none()
    {
        SeedDraft(newBoard: false);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab();

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(tab.ComponentPaneForTests.CreditsForTests);

            tab.CloseComponent();
        });

        Assert.Empty(Credits());
    }
}
