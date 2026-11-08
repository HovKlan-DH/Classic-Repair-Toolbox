using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A board's first draft, made when the Contribute tab's first change starts (Main.FirstDraft.cs -
// CRT 4.0 step 2, 2026-10-25; it replaced "Edit board as draft"). Driven on a real Main over a
// published board in a temp data tree; Main's StartAsync is never called (MainWindowTests' header).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainFirstDraftTests : IDisposable
{
    private const string ExcelDataFile = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private static readonly HardwareBoardEntry Entry = new()
    {
        HardwareName = "Commodore 64",
        BoardName = "250407",
        ExcelDataFile = MainFirstDraftTests.ExcelDataFile,
    };

    public MainFirstDraftTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));

        string dataRoot = this.thisWorkspace.Path_("Data");
        CachedWorkbooks.Write(
            DraftBoardSource.PublishedPathOf(dataRoot, MainFirstDraftTests.ExcelDataFile),
            new BoardData { Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU" }] });
        DataManager.LoadFrom(dataRoot, "does-not-exist.xlsx");
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static CRT.Main Window()
    {
        var window = new CRT.Main();

        window.TabDrafts.HardwareBoardsOverrideForTests = [MainFirstDraftTests.Entry];
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    private static string BoardFile => DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, MainFirstDraftTests.ExcelDataFile);

    // A board with no draft gets one - a copy of the published board - and the tab lists it.
    [Fact]
    public async Task A_board_with_no_draft_gets_a_copy_of_the_published_board()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainFirstDraftTests.Window();

            Assert.False(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, MainFirstDraftTests.ExcelDataFile));

            Assert.True(await window.EnsureDraftAsync(MainFirstDraftTests.Entry));

            BoardData draft = DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, MainFirstDraftTests.ExcelDataFile)!;
            Assert.Equal("U1", Assert.Single(draft.Components).BoardLabel);
            Assert.Single(window.TabDrafts.Drafts);
            Assert.Contains("hasDrafts", window.ContributeTabHeaderText.Classes);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** ONE DRAFT PER BOARD. *** With a draft already, nothing is written - the draft keeps every
    // change made to it - and the answer is simply yes: there is no "start a draft" to refuse.
    // ###########################################################################################
    [Fact]
    public async Task A_board_with_a_draft_keeps_it_and_nothing_is_written()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainFirstDraftTests.Window();
            Assert.True(await window.EnsureDraftAsync(MainFirstDraftTests.Entry));

            DateTime written = File.GetLastWriteTimeUtc(MainFirstDraftTests.BoardFile);

            Assert.True(await window.EnsureDraftAsync(MainFirstDraftTests.Entry));

            Assert.Equal(written, File.GetLastWriteTimeUtc(MainFirstDraftTests.BoardFile));
            Assert.Null(window.TabDrafts.DraftProblemForTests);

            window.Close();
        });
    }

    // No board, or one with no data file, has no draft to make.
    [Fact]
    public async Task No_board_has_no_draft_to_make()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainFirstDraftTests.Window();

            Assert.False(await window.EnsureDraftAsync(null));
            Assert.False(await window.EnsureDraftAsync(new HardwareBoardEntry { HardwareName = "X", BoardName = "Y", ExcelDataFile = " " }));

            window.Close();
        });
    }
}
