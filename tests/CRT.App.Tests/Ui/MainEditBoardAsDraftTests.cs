using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Edit board as draft" on the Contribute tab (owner request, 2026-10-09: "should create a new
// draft of the selected board, if it does not already exists. If it exits already, it should be
// informed via popup, that there cannot be two draft for same board") - Main.EditBoardAsDraft.cs,
// driven on a real Main over a published board in a temp data tree. Main's StartAsync is never
// called (MainWindowTests' header), and DraftExistsWindow's answer comes from the test.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainEditBoardAsDraftTests : IDisposable
{
    private const string ExcelDataFile = "Commodore/C64/250407/Data C64 250407.xlsx";

    private readonly TempWorkspace thisWorkspace = new();

    private static readonly HardwareBoardEntry Entry = new()
    {
        HardwareName = "Commodore 64",
        BoardName = "250407",
        ExcelDataFile = MainEditBoardAsDraftTests.ExcelDataFile,
    };

    public MainEditBoardAsDraftTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));

        string dataRoot = this.thisWorkspace.Path_("Data");
        CachedWorkbooks.Write(
            Path.Combine(dataRoot, MainEditBoardAsDraftTests.ExcelDataFile),
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
        var window = new CRT.Main
        {
            EditAsDraftBoardOverrideForTests = MainEditBoardAsDraftTests.Entry,
        };

        window.TabDrafts.HardwareBoardsOverrideForTests = [MainEditBoardAsDraftTests.Entry];
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    private static bool HasDraft() => DraftBoardSource.HasDraft(DraftManager.DraftsRoot, MainEditBoardAsDraftTests.ExcelDataFile);

    // A board with no draft gets one - a copy of the published board - and the contributor lands on
    // the Drafts tab with that draft open in the table.
    [Fact]
    public async Task A_board_with_no_draft_gets_one_and_its_table_opens_on_the_Drafts_tab()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainEditBoardAsDraftTests.Window();
            int asked = 0;
            window.ExistingDraftAnswerForTests = _ => { asked++; return false; };

            Assert.False(MainEditBoardAsDraftTests.HasDraft());

            await window.EditBoardAsDraftAsync();
            Dispatcher.UIThread.RunJobs();

            Assert.True(MainEditBoardAsDraftTests.HasDraft());
            Assert.Equal(0, asked);
            Assert.True(window.DraftsTabItem.IsVisible);
            Assert.Same(window.DraftsTabItem, window.MainTabControl.SelectedItem);
            Assert.True(window.TabDrafts.IsTableOpen);

            BoardData draft = DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, MainEditBoardAsDraftTests.ExcelDataFile)!;
            Assert.Equal("U1", Assert.Single(draft.Components).BoardLabel);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** ONE DRAFT PER BOARD. *** With a draft already, nothing is written - the draft keeps every
    // change made to it - and the window says so, naming the board. Closed, nothing else happens;
    // "Open the draft" lands on the same table.
    // ###########################################################################################
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task A_board_with_a_draft_is_offered_that_draft_and_nothing_is_written(bool openIt)
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainEditBoardAsDraftTests.Window();
            window.ExistingDraftAnswerForTests = _ => false;
            await window.EditBoardAsDraftAsync();
            window.MainTabControl.SelectedIndex = 0;
            await window.TabDrafts.CloseTableAsync();

            string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, MainEditBoardAsDraftTests.ExcelDataFile);
            DateTime written = File.GetLastWriteTimeUtc(workbook);
            var named = new List<string>();
            window.ExistingDraftAnswerForTests = name => { named.Add(name); return openIt; };

            await window.EditBoardAsDraftAsync();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Commodore 64 / 250407"], named);
            Assert.Equal(written, File.GetLastWriteTimeUtc(workbook));
            Assert.Equal(openIt, ReferenceEquals(window.MainTabControl.SelectedItem, window.DraftsTabItem));
            Assert.Equal(openIt, window.TabDrafts.IsTableOpen);

            window.Close();
        });
    }
}
