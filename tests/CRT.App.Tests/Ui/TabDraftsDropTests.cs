using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Files dropped on the Contribute tab (TabDrafts.Drop.cs - CRT 4.0 step 3, slice 3e): pictures of the
// board go into the Pictures part as its "Add pictures" adds them (ComponentPaneTests), its KiCad
// project folder opens "KiCad data" importing it (BoardFilesWindowTests); nothing usable is said on
// the tab's line.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsDropTests : IDisposable
{
    private const string Board = "Commodore/C64/999001/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsDropTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    // The tab on a board, the windows it would open recorded instead.
    private static (TabDrafts Tab, List<(BoardDropKind Kind, IReadOnlyList<string> Paths)> Opened) TabOnBoard()
    {
        var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "999001", ExcelDataFile = Board, IsDraftOnly = true };
        var opened = new List<(BoardDropKind, IReadOnlyList<string>)>();

        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [entry],
            PublishedBoardOverrideForTests = _ => null,
            EnsureDraftOverrideForTests = _ => Task.FromResult(true),
            ManageDroppedFilesOverrideForTests = (_, kind, paths) =>
            {
                opened.Add((kind, paths));
                return Task.CompletedTask;
            },
        };

        tab.RefreshDrafts();
        tab.ShowBoard(entry, new BoardData(), "PAL", null);
        return (tab, opened);
    }

    // The pictures alone are handed on - what is no picture left out, and said on the tab's line.
    [Fact]
    public async Task Pictures_dropped_are_handed_on_as_pictures_and_the_rest_left_out()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");
        string notes = this.thisWorkspace.WriteFile("notes.txt", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var opened) = TabOnBoard();

            await tab.DropOnBoardAsync([top, notes]);

            (BoardDropKind kind, IReadOnlyList<string> paths) = Assert.Single(opened);
            Assert.Equal(BoardDropKind.Pictures, kind);
            Assert.Equal([top], paths);
            Assert.Equal(BoardDrop.LeftWithoutPictures(1), tab.DraftProblemForTests);
        });
    }

    // ###########################################################################################
    // Pictures dropped on the board go into its Pictures part, opened for them, as one step - the
    // "Schematic images" window they opened until step 3 is gone.
    // ###########################################################################################
    [Fact]
    public async Task Pictures_dropped_are_added_in_the_boards_Pictures_part()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, Board);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }] });
        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, Board), new DraftMarker { BoardKey = Board });
        string under = this.thisWorkspace.WriteFile("Under side.png", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = TabOnBoard();
            tab.ManageDroppedFilesOverrideForTests = null;

            await tab.DropOnBoardAsync([under]);

            Assert.True(tab.IsComponentOpen);
            Assert.Equal(PaneBoardPart.Pictures, tab.ComponentPaneForTests.BoardPart);
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Board)!;
            Assert.Equal(["Top", "Under side"], session.Content.Board.Schematics.Select(row => row.SchematicName));
            Assert.Equal("Picture added", session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);
            Assert.Equal(["Top", "Under side"], tab.ComponentPaneForTests.BoardPicturesForTests.Select(row => row.Name));
        });
    }

    // ###########################################################################################
    // What was left out is said ONCE THE DRAFT IS THERE (code review, 2026-10-28): making a board's
    // first draft clears the tab's line (Main.EnsureDraftAsync - stood in for here as it does it), and
    // the note said before went unseen on every board with no draft yet.
    // ###########################################################################################
    [Fact]
    public async Task What_was_left_out_is_said_after_the_first_draft_is_made()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");
        string notes = this.thisWorkspace.WriteFile("notes.txt", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var opened) = TabOnBoard();
            tab.EnsureDraftOverrideForTests = _ =>
            {
                tab.ShowDraftProblem(null);
                return Task.FromResult(true);
            };

            await tab.DropOnBoardAsync([top, notes]);

            Assert.Single(opened);
            Assert.Equal(BoardDrop.LeftWithoutPictures(1), tab.DraftProblemForTests);
        });
    }

    // The plan's "drop the project folder": "KiCad data" opens importing it.
    [Fact]
    public async Task A_KiCad_project_folder_dropped_opens_KiCad_data_with_it()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Open128", "Open128.kicad_pro"), "{}");
        string project = Path.Combine(this.thisWorkspace.Root, "Open128");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var opened) = TabOnBoard();

            await tab.DropOnBoardAsync([project]);

            (BoardDropKind kind, IReadOnlyList<string> paths) = Assert.Single(opened);
            Assert.Equal(BoardDropKind.KiCad, kind);
            Assert.Equal([project], paths);
        });
    }

    // Nothing usable: said on the tab's line, and no window opens.
    [Fact]
    public async Task Nothing_usable_dropped_is_said_and_opens_nothing()
    {
        string pdf = this.thisWorkspace.WriteFile("datasheet.pdf", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var opened) = TabOnBoard();

            await tab.DropOnBoardAsync([pdf]);

            Assert.Empty(opened);
            Assert.Equal(BoardDrop.NothingUsable, tab.DraftProblemForTests);
        });
    }

    // ###########################################################################################
    // "CRT has to be updated" covers the tab, and nothing is done in it (code review, 2026-10-28): the
    // cover lies inside the tab, so a drop on it still reached the tab and opened a window over it.
    // ###########################################################################################
    [Fact]
    public async Task A_drop_on_a_tab_covered_by_CRT_has_to_be_updated_does_nothing()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var opened) = TabOnBoard();
            tab.ShowUpdateRequired(Handlers.Online.AppUpdateRequiredWording.For(Handlers.Online.AppUpdateArea.Drafts, "Please update CRT.", null));

            await tab.DropOnBoardAsync([top]);

            Assert.Empty(opened);
        });
    }

    // ###########################################################################################
    // Another board shown while what was dropped is read (code review, 2026-10-28): the files were
    // dropped on the board before - they never go into the one shown now.
    // ###########################################################################################
    [Fact]
    public async Task Another_board_shown_while_the_drop_is_read_takes_nothing()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var opened) = TabOnBoard();
            var reading = new TaskCompletionSource<BoardDrop>();
            tab.ReadDropOverrideForTests = _ => reading.Task;

            Task dropping = tab.DropOnBoardAsync([top]);

            var other = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "999002", ExcelDataFile = "Commodore/C64/999002/board.json", IsDraftOnly = true };
            tab.ShowBoard(other, new BoardData(), "PAL", null);
            reading.SetResult(new BoardDrop(BoardDropKind.Pictures, [top]));
            await dropping;

            Assert.Empty(opened);
        });
    }

    // No board shown: nothing to add to.
    [Fact]
    public async Task With_no_board_shown_a_drop_does_nothing()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");

        await UiTest.RunAsync(async () =>
        {
            var opened = new List<BoardDropKind>();
            var tab = new TabDrafts
            {
                ManageDroppedFilesOverrideForTests = (_, kind, _) =>
                {
                    opened.Add(kind);
                    return Task.CompletedTask;
                },
            };

            await tab.DropOnBoardAsync([top]);

            Assert.Empty(opened);
            Assert.Null(tab.DraftProblemForTests);
        });
    }
}
