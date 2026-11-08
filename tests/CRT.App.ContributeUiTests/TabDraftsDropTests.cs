using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Files dropped on the Contribute tab (TabDrafts.Drop.cs - CRT 4.0 step 3, slice 3e). *** ONLY "BOARD:
// SCHEMATIC IMAGES" TAKES THEM (owner decision, 2026-11-03: "it should not allow to drag any schematics"
// into the Overview page - "it is not the logical place to do this"). *** Images go in as "Add schematic
// image(s)" adds them (ComponentPaneTests); a KiCad project has its own drop target on "KiCad: Files"
// (KiCadDataPanelTests), and anything else is said on the tab's line.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsDropTests : IDisposable
{
    private const string Board = "Commodore/C64/999001/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsDropTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, Board);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }] });
        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, Board), new DraftMarker { BoardKey = Board });
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    // The tab on a board, what it would add recorded instead - on the Overview page, nothing open.
    private static (TabDrafts Tab, List<IReadOnlyList<string>> Added) TabOnBoard()
    {
        var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "999001", ExcelDataFile = Board, IsDraftOnly = true };
        var added = new List<IReadOnlyList<string>>();

        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [entry],
            PublishedBoardOverrideForTests = _ => null,
            EnsureDraftOverrideForTests = _ => Task.FromResult(true),
            AddDroppedPicturesOverrideForTests = (_, paths) =>
            {
                added.Add(paths);
                return Task.CompletedTask;
            },
        };

        tab.RefreshDrafts();
        tab.ShowBoard(entry, new BoardData(), "PAL", null);
        return (tab, added);
    }

    // The same, with "Board: Schematic images" open.
    private static async Task<(TabDrafts Tab, List<IReadOnlyList<string>> Added)> TabOnPicturesPage()
    {
        (TabDrafts tab, List<IReadOnlyList<string>> added) = TabOnBoard();

        await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
        Assert.Equal(PaneBoardPart.Pictures, tab.ComponentPaneForTests.BoardPart);

        return (tab, added);
    }

    // ###########################################################################################
    // The owner's report: images dropped on the Overview page went into the board. Nothing happens there,
    // nor on any other page - not even what was dropped is read.
    // ###########################################################################################
    [Theory]
    [InlineData(null)]
    [InlineData(nameof(PaneBoardPart.Files))]
    [InlineData(nameof(PaneBoardPart.Links))]
    public async Task Only_the_schematic_images_page_takes_what_is_dropped(string? page)
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var added) = TabOnBoard();

            if (page is not null)
            {
                await tab.OpenBoardPartAsync(Enum.Parse<PaneBoardPart>(page));
            }

            bool read = false;
            tab.ReadDropOverrideForTests = _ =>
            {
                read = true;
                return Task.FromResult(new BoardDrop(BoardDropKind.Pictures, [top]));
            };

            await tab.DropOnBoardAsync([top]);

            Assert.False(read);
            Assert.Empty(added);
            Assert.Null(tab.DraftProblemForTests);
        });
    }

    // The pictures alone are handed on - what is no picture left out, and said on the tab's line.
    [Fact]
    public async Task Pictures_dropped_are_handed_on_and_the_rest_left_out()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");
        string notes = this.thisWorkspace.WriteFile("notes.txt", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var added) = await TabOnPicturesPage();

            await tab.DropOnBoardAsync([top, notes]);

            Assert.Equal([top], Assert.Single(added));
            Assert.Equal(BoardDrop.LeftWithoutPictures(1), tab.DraftProblemForTests);
        });
    }

    // ###########################################################################################
    // Pictures dropped on the page go in as one step, as its "Add schematic image(s)" adds them.
    // ###########################################################################################
    [Fact]
    public async Task Pictures_dropped_are_added_on_the_page_as_one_step()
    {
        string under = this.thisWorkspace.WriteFile("Under side.png", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, _) = await TabOnPicturesPage();
            tab.AddDroppedPicturesOverrideForTests = null;

            await tab.DropOnBoardAsync([under]);

            Assert.True(tab.IsComponentOpen);
            Assert.Equal(PaneBoardPart.Pictures, tab.ComponentPaneForTests.BoardPart);
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Board)!;
            Assert.Equal(["Top", "Under side"], session.Content.Board.Schematics.Select(row => row.SchematicName));
            Assert.Equal("Schematic image added", session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);
            Assert.Equal(["Top", "Under side"], tab.ComponentPaneForTests.BoardPicturesForTests.Select(row => row.Name));
        });
    }

    // A KiCad project has its own page: told where it goes, and nothing is added.
    [Fact]
    public async Task A_KiCad_project_folder_dropped_is_told_where_it_goes()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Open128", "Open128.kicad_pro"), "{}");
        string project = Path.Combine(this.thisWorkspace.Root, "Open128");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var added) = await TabOnPicturesPage();

            await tab.DropOnBoardAsync([project]);

            Assert.Empty(added);
            Assert.False(tab.IsKiCadOpen);
            Assert.Equal(BoardDrop.KiCadGoesElsewhere, tab.DraftProblemForTests);
        });
    }

    // Nothing usable: said on the tab's line.
    [Fact]
    public async Task Nothing_usable_dropped_is_said_and_adds_nothing()
    {
        string pdf = this.thisWorkspace.WriteFile("datasheet.pdf", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var added) = await TabOnPicturesPage();

            await tab.DropOnBoardAsync([pdf]);

            Assert.Empty(added);
            Assert.Equal(BoardDrop.NothingUsable, tab.DraftProblemForTests);
        });
    }

    // ###########################################################################################
    // "CRT has to be updated" covers the tab, and nothing is done in it (code review, 2026-10-28): the
    // cover lies inside the tab, so a drop on it still reached the tab.
    // ###########################################################################################
    [Fact]
    public async Task A_drop_on_a_tab_covered_by_CRT_has_to_be_updated_does_nothing()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var added) = await TabOnPicturesPage();
            tab.ShowUpdateRequired(Handlers.Online.AppUpdateRequiredWording.For(Handlers.Online.AppUpdateArea.Drafts, "Please update CRT.", null));

            await tab.DropOnBoardAsync([top]);

            Assert.Empty(added);
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
            (TabDrafts tab, var added) = await TabOnPicturesPage();
            var reading = new TaskCompletionSource<BoardDrop>();
            tab.ReadDropOverrideForTests = _ => reading.Task;

            Task dropping = tab.DropOnBoardAsync([top]);

            var other = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "999002", ExcelDataFile = "Commodore/C64/999002/board.json", IsDraftOnly = true };
            tab.ShowBoard(other, new BoardData(), "PAL", null);
            reading.SetResult(new BoardDrop(BoardDropKind.Pictures, [top]));
            await dropping;

            Assert.Empty(added);
        });
    }

    // The page left while what was dropped is read: nothing goes in out of sight.
    [Fact]
    public async Task The_page_left_while_the_drop_is_read_takes_nothing()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var added) = await TabOnPicturesPage();
            var reading = new TaskCompletionSource<BoardDrop>();
            tab.ReadDropOverrideForTests = _ => reading.Task;

            Task dropping = tab.DropOnBoardAsync([top]);

            await tab.OpenBoardPartAsync(PaneBoardPart.Links);
            reading.SetResult(new BoardDrop(BoardDropKind.Pictures, [top]));
            await dropping;

            Assert.Empty(added);
        });
    }

    // No board shown: nothing to add to.
    [Fact]
    public async Task With_no_board_shown_a_drop_does_nothing()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");

        await UiTest.RunAsync(async () =>
        {
            var added = new List<IReadOnlyList<string>>();
            var tab = new TabDrafts
            {
                AddDroppedPicturesOverrideForTests = (_, paths) =>
                {
                    added.Add(paths);
                    return Task.CompletedTask;
                },
            };

            await tab.DropOnBoardAsync([top]);

            Assert.Empty(added);
            Assert.Null(tab.DraftProblemForTests);
        });
    }
}
