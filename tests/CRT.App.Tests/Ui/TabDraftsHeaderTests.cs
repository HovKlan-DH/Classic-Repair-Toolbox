using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE DRAFT'S HEADER, ON ITS ROW (TabDrafts.Header.cs - CRT 4.0 step 1's last slice, 2026-10-25):
// "Draft: 3 changes - Saved [Undo] [Redo]", Try again after a save that could not be made, Review
// changes with its "Undo changes to <component>", and the summary the Submit and Discard draft
// dialogs show.
//
// The row's Undo and Redo are the draft's SESSION's - the one history every editor shares - and a
// step taken from the row with the table closed is saved at once (nothing else would save it).
// With the table open, they are the table's own Ctrl+Z.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsHeaderTests : IDisposable
{
    private const string BoardA = "Commodore/C64/250407/Data.xlsx";
    private const string BoardB = "Commodore/C64/250425/Data.xlsx";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsHeaderTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry Entry(string board, string excelDataFile) => new()
    {
        HardwareName = "Commodore 64",
        BoardName = board,
        ExcelDataFile = excelDataFile,
    };

    // A draft of one component, U0, with a category and no name.
    private static void WriteDraft(string excelDataFile, string friendlyName = "")
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U0", Category = "IC", FriendlyName = friendlyName });

        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker { BoardKey = excelDataFile });
    }

    // Published: U0 named "CPU".
    private static BoardData Published()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U0", Category = "IC", FriendlyName = "CPU" });
        return board;
    }

    private (TabDrafts Tab, List<string> Reloaded) BuildTab(params HardwareBoardEntry[] entries)
    {
        var reloaded = new List<string>();

        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [.. entries],
            PublishedBoardOverrideForTests = _ => TabDraftsHeaderTests.Published(),
            ReloadBoardOverrideForTests = reloaded.Add,
        };

        tab.RefreshDrafts();
        return (tab, reloaded);
    }

    // U0 renamed - these drafts hold nothing but components.
    private static Func<BoardContent, BoardContent?> Named(string friendlyName) =>
        content => content with
        {
            Board = new BoardData
            {
                RevisionDate = content.Board.RevisionDate,
                Components = [.. content.Board.Components.Select(component =>
                    component.BoardLabel == "U0" ? new ComponentEntry { BoardLabel = "U0", Category = "IC", FriendlyName = friendlyName } : component)],
            },
        };

    private static string NameOnDisk(string excelDataFile) =>
        DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, excelDataFile)!.Components[0].FriendlyName;

    private static DraftSession SessionOf(string excelDataFile) =>
        DraftManager.Sessions.Get(DraftManager.DraftsRoot, excelDataFile)!;

    // ###########################################################################################
    // With the table closed, the row's Undo takes the session's last step - whichever editor made
    // it - SAVES it, and the board on screen catches up. Redo puts it back the same way.
    // ###########################################################################################
    [Fact]
    public async Task The_rows_undo_takes_the_sessions_step_and_saves_it_at_once()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", TabDraftsHeaderTests.BoardA);
            (TabDrafts tab, List<string> reloaded) = this.BuildTab(entry);

            DraftSession session = SessionOf(TabDraftsHeaderTests.BoardA);
            session.Apply("Changes to [U0]", TabDraftsHeaderTests.Named("ULA"));
            Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
            tab.RefreshDrafts();

            DraftHeaderState header = tab.HeaderForTests(TabDraftsHeaderTests.BoardA)!;
            Assert.True(header.CanUndo);
            Assert.Equal("Undo: Changes to [U0]", header.UndoTip);

            await tab.UndoAsync(entry);

            Assert.Equal(string.Empty, NameOnDisk(TabDraftsHeaderTests.BoardA));
            Assert.Equal(DraftSaveWording.Saved, header.SaveStateText);
            Assert.True(header.CanRedo);
            Assert.Equal([TabDraftsHeaderTests.BoardA], reloaded);

            await tab.RedoAsync(entry);

            Assert.Equal("ULA", NameOnDisk(TabDraftsHeaderTests.BoardA));
            Assert.True(header.CanUndo);
        });
    }

    // With the table open, Undo is the table's own - it hands over what it holds and undoes that,
    // and the row does not save or reload behind it.
    [Fact]
    public async Task With_the_table_open_the_rows_undo_is_the_tables()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", TabDraftsHeaderTests.BoardA);
            (TabDrafts tab, List<string> reloaded) = this.BuildTab(entry);
            await tab.OpenTableAsync(entry);

            BoardTableSheet components = tab.TableEditor.DocumentForTests!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
            int friendlyName = components.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
            components.Rows[0].Cells[friendlyName].Text = "Typed in the table";

            await tab.UndoAsync(entry);

            DraftSession session = SessionOf(TabDraftsHeaderTests.BoardA);
            Assert.Equal(string.Empty, session.Content.Board.Components[0].FriendlyName);
            Assert.Equal(
                string.Empty,
                tab.TableEditor.DocumentForTests!.FindSheet(BoardWorkbookSchema.SheetComponents)!.Rows[0].Cells[friendlyName].Text);
            Assert.True(session.CanRedo);
            Assert.Empty(reloaded);
        });
    }

    // ###########################################################################################
    // A header is made before any editor opens a session (the draft is just Saved), follows the one
    // opened later, and stays the SAME object through every rebuilt row - compared by reference, so
    // a kept row stays the same row.
    // ###########################################################################################
    [Fact]
    public void The_header_follows_a_session_opened_later_and_survives_a_rebuilt_row()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA);

        UiTest.Run(() =>
        {
            (TabDrafts tab, _) = this.BuildTab(Entry("250407", TabDraftsHeaderTests.BoardA));

            DraftHeaderState header = tab.HeaderForTests(TabDraftsHeaderTests.BoardA)!;
            Assert.Null(header.SessionForTests);
            Assert.Same(header, Assert.Single(tab.Drafts).Header);

            DraftSession session = SessionOf(TabDraftsHeaderTests.BoardA);
            tab.RefreshDrafts();

            Assert.Same(header, tab.HeaderForTests(TabDraftsHeaderTests.BoardA));
            Assert.Same(session, header.SessionForTests);
            Assert.Same(header, Assert.Single(tab.Drafts).Header);
        });
    }

    // A draft no longer listed lets go of its session - its header is gone.
    [Fact]
    public void A_draft_no_longer_listed_lets_go_of_its_header()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA);
        WriteDraft(TabDraftsHeaderTests.BoardB);

        UiTest.Run(() =>
        {
            (TabDrafts tab, _) = this.BuildTab(Entry("250407", TabDraftsHeaderTests.BoardA), Entry("250425", TabDraftsHeaderTests.BoardB));
            DraftSession session = SessionOf(TabDraftsHeaderTests.BoardB);
            tab.RefreshDrafts();
            DraftHeaderState header = tab.HeaderForTests(TabDraftsHeaderTests.BoardB)!;
            Assert.Same(session, header.SessionForTests);

            tab.HardwareBoardsOverrideForTests = [Entry("250407", TabDraftsHeaderTests.BoardA)];
            tab.RefreshDrafts();

            Assert.Null(tab.HeaderForTests(TabDraftsHeaderTests.BoardB));
            Assert.Null(header.SessionForTests);
            Assert.NotNull(tab.HeaderForTests(TabDraftsHeaderTests.BoardA));
        });
    }

    // Try again with the table closed saves what waits.
    [Fact]
    public async Task Try_again_saves_what_waits()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", TabDraftsHeaderTests.BoardA);
            (TabDrafts tab, _) = this.BuildTab(entry);

            SessionOf(TabDraftsHeaderTests.BoardA).Apply("Changes to [U0]", TabDraftsHeaderTests.Named("ULA"));

            await tab.TryAgainAsync(entry);

            Assert.Equal("ULA", NameOnDisk(TabDraftsHeaderTests.BoardA));
            Assert.Equal(DraftSaveState.Saved, SessionOf(TabDraftsHeaderTests.BoardA).SaveState);
        });
    }

    // ###########################################################################################
    // *** A FILE CHANGED SOMEWHERE ELSE IS NOT TRIED AGAIN FROM THE ROW. *** It would be refused
    // again, for ever. The table opens on it instead, and its bar offers both versions - nothing
    // written over and nothing thrown away unasked.
    // ###########################################################################################
    [Fact]
    public async Task Try_again_after_a_change_somewhere_else_opens_the_table_offering_both_versions()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", TabDraftsHeaderTests.BoardA);
            (TabDrafts tab, _) = this.BuildTab(entry);

            DraftSession session = SessionOf(TabDraftsHeaderTests.BoardA);
            tab.RefreshDrafts();
            session.Apply("Changes to [U0]", TabDraftsHeaderTests.Named("mine"));
            WriteDraft(TabDraftsHeaderTests.BoardA, friendlyName: "theirs");
            Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.Save());
            Assert.True(tab.HeaderForTests(TabDraftsHeaderTests.BoardA)!.IsFailed);

            await tab.TryAgainAsync(entry);

            Assert.True(tab.IsTableOpen);
            Assert.True(tab.TableEditor.GetControl<Border>("ChangedOnDiskBar").IsVisible);
            Assert.Equal("theirs", NameOnDisk(TabDraftsHeaderTests.BoardA));
            Assert.Equal("mine", session.Content.Board.Components[0].FriendlyName);
        });
    }

    // ###########################################################################################
    // The summary is of the draft AS EVERY EDITOR SHOWS IT - a change its session holds unsaved is
    // in it - against the published board.
    // ###########################################################################################
    [Fact]
    public async Task The_summary_reads_the_draft_held_in_memory_against_the_published_board()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA, friendlyName: "CPU");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", TabDraftsHeaderTests.BoardA);
            (TabDrafts tab, _) = this.BuildTab(entry);

            Assert.True((await tab.SummaryForAsync(entry)).IsEmpty);

            SessionOf(TabDraftsHeaderTests.BoardA).Apply("Changes to [U0]", TabDraftsHeaderTests.Named("ULA"));

            DraftChangeSummary summary = await tab.SummaryForAsync(entry);

            DraftChangeGroup group = Assert.Single(summary.Groups);
            Assert.Equal("U0", group.Label);
            Assert.Equal(DraftComponentChange.Changed, group.Change);
        });
    }

    // ###########################################################################################
    // "Undo changes to [U0]": the component back to its published values as ONE step - saved, the
    // board on screen caught up, the summary read again without it - and the row's Undo brings the
    // changes back.
    // ###########################################################################################
    [Fact]
    public async Task Undo_changes_to_a_component_puts_it_back_as_one_step_the_rows_undo_takes_back()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA, friendlyName: "ULA");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", TabDraftsHeaderTests.BoardA);
            (TabDrafts tab, List<string> reloaded) = this.BuildTab(entry);

            (DraftChangeSummary summary, string? failure) = await tab.UndoComponentChangesAsync(entry, "U0");

            Assert.Null(failure);
            Assert.True(summary.IsEmpty);
            Assert.Equal("CPU", NameOnDisk(TabDraftsHeaderTests.BoardA));
            Assert.Equal([TabDraftsHeaderTests.BoardA], reloaded);

            DraftHeaderState header = tab.HeaderForTests(TabDraftsHeaderTests.BoardA)!;
            Assert.Equal("Undo: " + DraftStepWording.ComponentReverted("U0"), header.UndoTip);

            await tab.UndoAsync(entry);

            Assert.Equal("ULA", NameOnDisk(TabDraftsHeaderTests.BoardA));
        });
    }

    // ###########################################################################################
    // The row's Undo, Redo, Try again and Review changes are bound to its own commands - a button
    // left on another command, or none, compiles and does nothing. Try again shows only after a save
    // that could not be made.
    // ###########################################################################################
    [Fact]
    public void The_rows_header_buttons_are_bound_to_its_commands()
    {
        WriteDraft(TabDraftsHeaderTests.BoardA);

        UiTest.Run(() =>
        {
            HardwareBoardEntry entry = Entry("250407", TabDraftsHeaderTests.BoardA);
            (TabDrafts tab, _) = this.BuildTab(entry);

            // The header is the board shown's (CRT 4.0 step 2).
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            var window = new Window { Content = tab, Width = 1300, Height = 600 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            List<Button> buttons = tab.GetVisualDescendants().OfType<Button>().ToList();
            Button Named(string content) => Assert.Single(buttons, button => (button.Content as string) == content);

            DraftListItem row = Assert.Single(tab.Drafts);

            Assert.Same(row.UndoCommand, Named("Undo").Command);
            Assert.Same(row.RedoCommand, Named("Redo").Command);
            Assert.Same(row.TryAgainCommand, Named(DraftSaveWording.TryAgain).Command);
            Assert.Same(row.ReviewChangesCommand, Named("Review changes").Command);

            Assert.False(Named("Undo").IsEnabled);
            Assert.False(Named(DraftSaveWording.TryAgain).IsVisible);

            window.Close();
        });
    }
}
