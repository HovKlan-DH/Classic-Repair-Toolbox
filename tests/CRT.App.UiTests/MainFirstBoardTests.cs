using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// *** THE SPLASH WAITS FOR THE FIRST BOARD (owner report, 2026-11-09: "It can take a couple of seconds before it will
// populate the schematic image and thumbnails etc. and in the meantime I see an empty UI"). *** App shows Main only
// once Main.FirstBoardShown is done (Main.FirstBoard.cs) - so these hold that it is done with the board really
// drawn, in a window NOT shown yet, and that it is done whatever else came of the load, so the splash never waits
// out its whole limit for a board that will not come.
//
// As MainBoardSwitchTests, through the real Main with real board files, and StartAsync never called (no board view is
// counted - one would be sent to the live service).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainFirstBoardTests : IDisposable
{
    private const string Hardware = "Commodore 64";

    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    private readonly TempWorkspace thisWorkspace = new();

    public MainFirstBoardTests()
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

    private string DataRoot => this.thisWorkspace.Path_("Data");

    // The data root listing `board` - with a picture, a component, both, or no board file at all.
    private void ListBoard(string board, bool withPicture = true, bool withFile = true)
    {
        var listed = new MainFiles.Board("Commodore/C64/" + board, MainFirstBoardTests.Hardware, board);
        DataManager.LoadFrom(this.DataRoot, MainFiles.Write(this.DataRoot, listed));

        if (!withFile)
            return;

        BoardDocument.StoredFileRow picture = StoredFiles.Put(this.DataRoot, "Images/Top.png", MainFirstBoardTests.Png);

        StoredFiles.WriteBoard(
            DraftBoardSource.PublishedPathOf(this.DataRoot, listed.Identity),
            new BoardData
            {
                Schematics = withPicture ? [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = picture.Name }] : [],
                Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
            },
            picture);
    }

    private static async Task UntilFirstBoardShown(CRT.Main main)
    {
        for (int attempt = 0; attempt < 1000 && !main.FirstBoardShown.IsCompleted; attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(main.FirstBoardShown.IsCompleted, "The first board load never said it was over - the splash would wait out its whole limit");
    }

    [Fact]
    public async Task The_first_board_is_drawn_in_full_before_the_window_is_shown()
    {
        this.ListBoard("250407");

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();

            try
            {
                // As StartAsync starts it, and the window NOT shown - App shows it once this is done.
                main.PopulateHardwareDropDown();
                Assert.False(main.FirstBoardShown.IsCompleted);

                await MainFirstBoardTests.UntilFirstBoardShown(main);

                TabSchematics tab = main.TabSchematicsControl;
                Assert.False(main.IsVisible);
                Assert.Equal("Top", tab.SelectedThumbnail?.Name);
                Assert.NotNull(tab.SchematicsImage.Source);
                Assert.Equal(["U1"], main.ComponentFilterListBox.Items.OfType<ComponentListItem>().Select(item => item.BoardLabel));
                Assert.Equal("250407", main.BoardComboBox.SelectedItem);
                Assert.False(main.IsCountingBoardViewForTests);
            }
            finally
            {
                main.Close();
            }
        });
    }

    // Nothing to draw: over as soon as that is known, so the splash does not wait out its limit.
    [Fact]
    public async Task With_no_board_listed_the_first_board_is_over_at_once()
    {
        DataManager.LoadFrom(this.DataRoot, MainFiles.Write(this.DataRoot));

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();

            try
            {
                main.PopulateHardwareDropDown();
                Assert.True(main.FirstBoardShown.IsCompleted);
                Assert.True(await main.WaitForFirstBoardAsync(TimeSpan.FromSeconds(5)));
            }
            finally
            {
                main.Close();
            }
        });
    }

    [Fact]
    public async Task A_board_with_no_pictures_is_over_once_it_is_read()
    {
        this.ListBoard("250425", withPicture: false);

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();

            try
            {
                main.PopulateHardwareDropDown();
                await MainFirstBoardTests.UntilFirstBoardShown(main);

                Assert.Null(main.TabSchematicsControl.SelectedThumbnail);
                Assert.Equal(["U1"], main.ComponentFilterListBox.Items.OfType<ComponentListItem>().Select(item => item.BoardLabel));
            }
            finally
            {
                main.Close();
            }
        });
    }

    // Listed, but its board file is not there: the load ends with nothing to show, and says so.
    [Fact]
    public async Task A_board_whose_file_is_missing_is_over_all_the_same()
    {
        this.ListBoard("250466", withFile: false);

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();

            try
            {
                main.PopulateHardwareDropDown();
                await MainFirstBoardTests.UntilFirstBoardShown(main);

                Assert.Null(main.TabSchematicsControl.SelectedThumbnail);
            }
            finally
            {
                main.Close();
            }
        });
    }

    // The limit holds: a board that never comes keeps the window hidden only that long.
    [Fact]
    public async Task The_wait_for_the_first_board_gives_up_after_its_time()
    {
        this.ListBoard("250469");

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();

            try
            {
                // No load started at all.
                Assert.False(await main.WaitForFirstBoardAsync(TimeSpan.FromMilliseconds(50)));
            }
            finally
            {
                main.Close();
            }
        });
    }
}
