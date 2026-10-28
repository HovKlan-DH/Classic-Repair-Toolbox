using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Files dropped on the Contribute tab (TabDrafts.Drop.cs - CRT 4.0 step 3, slice 3e): pictures of the
// board open "Schematic images" with them being added, its KiCad project folder "KiCad data" - the
// windows' own imports, tested in BoardFilesWindowTests; nothing usable is said on the tab's line.
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
    private static (TabDrafts Tab, List<(BoardFilesSection Section, IReadOnlyList<string> Paths)> Opened) TabOnBoard()
    {
        var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "999001", ExcelDataFile = Board, IsDraftOnly = true };
        var opened = new List<(BoardFilesSection, IReadOnlyList<string>)>();

        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [entry],
            PublishedBoardOverrideForTests = _ => null,
            EnsureDraftOverrideForTests = _ => Task.FromResult(true),
            ManageDroppedFilesOverrideForTests = (_, section, paths) =>
            {
                opened.Add((section, paths));
                return Task.CompletedTask;
            },
        };

        tab.RefreshDrafts();
        tab.ShowBoard(entry, new BoardData(), "PAL", null);
        return (tab, opened);
    }

    [Fact]
    public async Task Pictures_dropped_open_Schematic_images_with_them()
    {
        string top = this.thisWorkspace.WriteFile("top.png", "x");
        string notes = this.thisWorkspace.WriteFile("notes.txt", "x");

        await UiTest.RunAsync(async () =>
        {
            (TabDrafts tab, var opened) = TabOnBoard();

            await tab.DropOnBoardAsync([top, notes]);

            (BoardFilesSection section, IReadOnlyList<string> paths) = Assert.Single(opened);
            Assert.Equal(BoardFilesSection.SchematicImages, section);
            Assert.Equal([top], paths);
            Assert.Null(tab.DraftProblemForTests);
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

            (BoardFilesSection section, IReadOnlyList<string> paths) = Assert.Single(opened);
            Assert.Equal(BoardFilesSection.KiCadData, section);
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
            var opened = new List<BoardFilesSection>();
            var tab = new TabDrafts
            {
                ManageDroppedFilesOverrideForTests = (_, section, _) =>
                {
                    opened.Add(section);
                    return Task.CompletedTask;
                },
            };

            await tab.DropOnBoardAsync([top]);

            Assert.Empty(opened);
            Assert.Null(tab.DraftProblemForTests);
        });
    }
}
