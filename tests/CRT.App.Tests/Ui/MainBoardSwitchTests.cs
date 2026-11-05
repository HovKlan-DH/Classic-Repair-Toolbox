using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// *** THE REPORTED BUG (a user of CRT 3.1.0, 2026-11-05: "did you remove all schematics for the
// 250466 board?"). *** Nothing had been removed. With a thumbnail selected and the Schematics tab
// showing, EVERY SECOND board picked in the drop-down was never loaded: emptying the old tiles told
// the tab first, the tab told the gallery nothing was selected, and the gallery set its list's
// selection while the list had not seen the tiles go. The list read the tile it let go of from the
// empty collection and threw, out of Clear(), and OnBoardSelectionChanged's catch only logged it.
// So the drop-down named the new board over the PREVIOUS board's tiles and components, with no
// picture. The next board worked, as nothing was selected any more.
//
// SchematicsThumbnailGalleryTests holds that step alone. This is the user's own action, through
// the real Main with real board files, because nothing drove it before and the bug shipped: pick a
// board, then another, then go back - each shows its own pictures and components.
//
// Main's StartAsync is never called (MainWindowTests' header), so Main counts no board view: one
// of a board these tests load would be sent to the live service. The test asserts that too, as
// these boards are published and would otherwise be counted.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainBoardSwitchTests : IDisposable
{
    private const string Hardware = "Commodore 64";

    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    private readonly TempWorkspace thisWorkspace = new();

    public MainBoardSwitchTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
    }

    public void Dispose()
    {
        PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        DraftManager.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    // A data root listing these boards, each with one picture and one component of its own.
    private void LoadBoards(params string[] boardNames)
    {
        string dataRoot = this.thisWorkspace.Path_("Data");

        MainFiles.Board[] listed = boardNames
            .Select(name => new MainFiles.Board("Commodore/C64/" + name, MainBoardSwitchTests.Hardware, name))
            .ToArray();

        DataManager.LoadFrom(dataRoot, MainFiles.Write(dataRoot, listed));

        foreach (MainFiles.Board board in listed)
        {
            BoardDocument.StoredFileRow picture = StoredFiles.Put(dataRoot, $"Images/{board.BoardName}.png", MainBoardSwitchTests.Png);

            StoredFiles.WriteBoard(
                DraftBoardSource.PublishedPathOf(dataRoot, board.Identity),
                new BoardData
                {
                    Schematics = [new BoardSchematicEntry { SchematicName = MainBoardSwitchTests.PictureOf(board.BoardName), SchematicImageFile = picture.Name }],
                    Components = [new ComponentEntry { BoardLabel = MainBoardSwitchTests.ComponentOf(board.BoardName), Category = "IC" }],
                },
                picture);
        }
    }

    private static string PictureOf(string board) => "Picture of " + board;

    private static string ComponentOf(string board) => "U-" + board;

    private static async Task Until(Func<bool> condition, string what)
    {
        for (int attempt = 0; attempt < 1000 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(condition(), what);
    }

    private static CRT.Main ShowMain()
    {
        var main = new CRT.Main();
        main.Show();

        // As StartAsync starts it: the first board load cascades from the hardware's selection.
        main.PopulateHardwareDropDown();
        Dispatcher.UIThread.RunJobs();

        return main;
    }

    // Every surface the user looks at, as the screenshot of the report showed them disagreeing.
    private static async Task AssertShowsAsync(CRT.Main main, string board)
    {
        TabSchematics tab = main.TabSchematicsControl;
        ListBox tiles = tab.InlineThumbnailGallery.GalleryListBox;

        List<string?> TilesOnScreen() => tiles.GetRealizedContainers()
            .Select(container => (container.DataContext as SchematicThumbnail)?.Name)
            .ToList();

        List<string> Components() => main.ComponentFilterListBox.Items
            .OfType<ComponentListItem>()
            .Select(item => item.BoardLabel)
            .ToList();

        await MainBoardSwitchTests.Until(
            () => tab.SelectedThumbnail?.Name == MainBoardSwitchTests.PictureOf(board)
                && tab.SchematicsImage.Source != null
                && TilesOnScreen().SequenceEqual([MainBoardSwitchTests.PictureOf(board)])
                && Components().SequenceEqual([MainBoardSwitchTests.ComponentOf(board)]),
            $"[{board}] is not what is shown: selected [{tab.SelectedThumbnail?.Name}], " +
            $"picture {(tab.SchematicsImage.Source == null ? "blank" : "shown")}, " +
            $"tiles [{string.Join("; ", TilesOnScreen())}], components [{string.Join("; ", Components())}]");

        Assert.Equal(board, main.BoardComboBox.SelectedItem);
    }

    [Fact]
    public async Task Every_board_picked_in_the_drop_down_shows_its_own_pictures_and_components()
    {
        this.LoadBoards("250407", "250425", "250466", "250469");

        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = MainBoardSwitchTests.ShowMain();

            try
            {
                await MainBoardSwitchTests.AssertShowsAsync(main, "250407");

                // Back to boards seen before too: 250407 decoded again, as Main keeps three boards'
                // thumbnails (ThumbnailSetsKept) and had let it go, and 250466 out of that cache.
                foreach (string board in new[] { "250425", "250466", "250469", "250407", "250466" })
                {
                    main.BoardComboBox.SelectedItem = board;
                    await MainBoardSwitchTests.AssertShowsAsync(main, board);
                }

                Assert.False(main.IsCountingBoardViewForTests);
            }
            finally
            {
                main.Close();
            }
        });
    }
}
