using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE WORKSPACE (TabDrafts.Workspace.cs - CRT 4.0 step 2, 2026-10-25: "One tab instead of
// Contribute and Drafts ... A My drafts list in the header switches between boards ... The table
// view is one view of the same draft").
//
// The tab is about the board shown in CRT's drop-down lists (ShowBoard): its draft's header - or
// what it is, with none - its parts and components on the left, and the table view on the right.
// Every action on the left works on the board's draft, made first when it has none.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsWorkspaceTests : IDisposable
{
    private const string BoardA = "Commodore/C64/250407/board.json";
    private const string BoardB = "Commodore/C64/250425/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsWorkspaceTests()
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

    private static void WriteDraft(string excelDataFile, BoardData? content = null)
    {
        var board = content ?? new BoardData();

        if (content is null)
        {
            board.Components.Add(new ComponentEntry { BoardLabel = "U0", Category = "IC", FriendlyName = "CPU" });
        }

        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker { BoardKey = excelDataFile });
    }

    // A board of three components - one PAL only, one NTSC only - with every part counted.
    private static BoardData Board() => new()
    {
        RevisionDate = "2026-October-4",
        Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "VIC", Region = "PAL" },
            new ComponentEntry { BoardLabel = "U2", Category = "IC", FriendlyName = "VIC", Region = "NTSC" },
            new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" },
        ],
        KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "CLK", KiCadNetName = "CLK" }],
        BoardLinks = [new BoardLinkEntry { Name = "Forum" }],
        Credits = [],
    };

    private static TabDrafts Tab(params HardwareBoardEntry[] entries)
    {
        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [.. entries],
            PublishedBoardOverrideForTests = _ => null,
        };

        tab.RefreshDrafts();
        return tab;
    }

    private static string Text(TabDrafts tab, string name) => tab.GetControl<TextBlock>(name).Text ?? string.Empty;

    // ###########################################################################################
    // The board shown, in the left column: its parts counted, its components in the region shown,
    // and the ones its draft changed marked.
    // ###########################################################################################
    [Fact]
    public void The_left_column_counts_the_boards_parts_and_lists_its_components_in_the_region_shown()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();
            tab.ShowBoard(Entry("250407", BoardA), Board(), "PAL", new HashSet<string>(["U8"], StringComparer.OrdinalIgnoreCase));

            Assert.Equal("1", Text(tab, "PicturesCountText"));
            Assert.Equal("1", Text(tab, "SignalsCountText"));
            Assert.Equal("1", Text(tab, "FilesCountText"));
            Assert.Equal("none", Text(tab, "CreditsCountText"));

            Assert.Equal(["U1", "U8"], tab.ComponentsShownForTests.Select(item => item.BoardLabel));
            Assert.True(tab.ComponentsShownForTests.Single(item => item.BoardLabel == "U8").IsDrafted);
            Assert.False(tab.ComponentsShownForTests.Single(item => item.BoardLabel == "U1").IsDrafted);
            Assert.Contains("2026-October-4", Text(tab, "PublishedOnText"), StringComparison.Ordinal);
        });
    }

    // The search box narrows the list by its words; one that matches nothing says so.
    [Fact]
    public void The_search_narrows_the_components_and_says_when_nothing_matches()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();
            tab.ShowBoard(Entry("250407", BoardA), Board(), "PAL", null);

            TextBox search = tab.GetControl<TextBox>("ComponentSearchBox");

            // TextChanged is raised through the dispatcher.
            search.Text = "cpu";
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["U8"], tab.ComponentsShownForTests.Select(item => item.BoardLabel));
            Assert.False(tab.GetControl<TextBlock>("NoComponentsText").IsVisible);

            search.Text = "zilog";
            Dispatcher.UIThread.RunJobs();
            Assert.Empty(tab.ComponentsShownForTests);
            Assert.True(tab.GetControl<TextBlock>("NoComponentsText").IsVisible);
            Assert.Equal("No component matches the search.", Text(tab, "NoComponentsText"));
        });
    }

    // ###########################################################################################
    // With no board shown nothing on the left can be used - but "Add a new board" can: it is what
    // someone does when the board they want is not in the lists at all.
    // ###########################################################################################
    [Fact]
    public void Without_a_board_the_left_column_is_off_and_Add_a_new_board_is_not()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();
            tab.ShowBoard(null, null, "PAL", null);

            Assert.False(tab.GetControl<Button>("AddComponentButton").IsEffectivelyEnabled);
            Assert.False(tab.GetControl<Button>("TableViewButton").IsEffectivelyEnabled);
            Assert.True(tab.GetControl<Button>("AddNewBoardButton").IsEffectivelyEnabled);
            Assert.Equal("No board shown", Text(tab, "NoDraftBoardText"));
        });
    }

    // Both "add" buttons are green - they go ahead and make something (CLAUDE.md, button colours).
    [Fact]
    public void Add_component_and_Add_a_new_board_are_green()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();

            Assert.Contains("Accept", tab.GetControl<Button>("AddComponentButton").Classes);
            Assert.Contains("Accept", tab.GetControl<Button>("AddNewBoardButton").Classes);
            Assert.DoesNotContain("Accept", tab.GetControl<Button>("TableViewButton").Classes);
        });
    }

    // ###########################################################################################
    // The header: a board with a draft shows the draft's row; one with none says what it is, and
    // that the first change makes the draft.
    // ###########################################################################################
    [Fact]
    public void The_header_is_the_drafts_row_or_says_the_board_has_no_draft_yet()
    {
        WriteDraft(BoardA);

        UiTest.Run(() =>
        {
            HardwareBoardEntry withDraft = Entry("250407", BoardA);
            HardwareBoardEntry withoutDraft = Entry("250425", BoardB);
            TabDrafts tab = Tab(withDraft, withoutDraft);

            tab.ShowBoard(withDraft, Board(), "PAL", null);

            ContentControl header = tab.GetControl<ContentControl>("CurrentDraftRow");
            Assert.True(header.IsVisible);
            Assert.Equal(BoardA, Assert.IsType<DraftListItem>(header.Content).ExcelDataFile);
            Assert.False(tab.GetControl<Border>("NoDraftPanel").IsVisible);

            tab.ShowBoard(withoutDraft, Board(), "PAL", null);

            Assert.False(header.IsVisible);
            Assert.True(tab.GetControl<Border>("NoDraftPanel").IsVisible);
            Assert.Equal("Commodore 64 - 250425", Text(tab, "NoDraftBoardText"));
            Assert.Equal(ContributeBoardParts.NoDraftLine(hasBoard: true), Text(tab, "NoDraftText"));
        });
    }

    // ###########################################################################################
    // "My drafts" shows the board picked - through Main's drop-downs, so all of CRT is on it - and
    // falls back to its placeholder, so it reads as a menu, never as a second board name.
    // ###########################################################################################
    [Fact]
    public void Picking_a_draft_in_My_drafts_shows_its_board_and_the_list_falls_back()
    {
        WriteDraft(BoardA);
        WriteDraft(BoardB);

        UiTest.Run(() =>
        {
            TabDrafts tab = Tab(Entry("250407", BoardA), Entry("250425", BoardB));
            var shown = new List<string>();
            tab.SelectBoardOverrideForTests = entry => shown.Add(entry.ExcelDataFile);

            ComboBox myDrafts = tab.GetControl<ComboBox>("MyDraftsComboBox");
            Assert.Equal("My drafts (2)", myDrafts.PlaceholderText);

            myDrafts.SelectedItem = tab.Drafts.Single(item => item.ExcelDataFile == BoardB);

            Assert.Equal([BoardB], shown);
            Assert.Null(myDrafts.SelectedItem);
        });
    }

    // ###########################################################################################
    // *** A FIRST CHANGE MAKES THE DRAFT. *** On a board with none, an action on the left asks for
    // the draft first (Main.EnsureDraftAsync) and goes on once there is one; when none could be made,
    // it does nothing.
    // ###########################################################################################
    [Fact]
    public async Task An_action_on_a_board_with_no_draft_makes_one_first_and_stops_when_none_is_made()
    {
        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            var asked = new List<string>();
            var acted = new List<string>();

            tab.EnsureDraftOverrideForTests = board =>
            {
                asked.Add(board.ExcelDataFile);
                return Task.FromResult(false);
            };

            await tab.OnCurrentDraftAsync(board => { acted.Add(board.ExcelDataFile); return Task.CompletedTask; });

            Assert.Equal([BoardA], asked);
            Assert.Empty(acted);

            tab.EnsureDraftOverrideForTests = board =>
            {
                asked.Add(board.ExcelDataFile);
                WriteDraft(board.ExcelDataFile);
                return Task.FromResult(true);
            };

            await tab.OnCurrentDraftAsync(board => { acted.Add(board.ExcelDataFile); return Task.CompletedTask; });

            Assert.Equal([BoardA, BoardA], asked);
            Assert.Equal([BoardA], acted);

            // Has one now: nothing more is asked.
            await tab.OnCurrentDraftAsync(board => { acted.Add(board.ExcelDataFile); return Task.CompletedTask; });

            Assert.Equal(2, asked.Count);
            Assert.Equal(2, acted.Count);
            Assert.Single(tab.Drafts);
        });
    }

    // A part of the board with no screen of its own yet opens the table view on its sheet.
    [Fact]
    public async Task Credits_open_the_table_view_on_its_sheet()
    {
        WriteDraft(BoardA);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.ShowTableOnSheetAsync(BoardWorkbookSchema.SheetCredits);

            Assert.True(tab.IsTableOpen);
            Assert.Equal(BoardWorkbookSchema.SheetCredits, tab.TableEditor.CurrentSheet!.Name);
            Assert.Equal("Close table view", tab.GetControl<Button>("TableViewButton").Content);

            await tab.ToggleTableViewAsync();

            Assert.False(tab.IsTableOpen);
            Assert.Equal("Table view", tab.GetControl<Button>("TableViewButton").Content);
        });
    }

    // ###########################################################################################
    // *** THE TABLE IS A VIEW OF THE BOARD SHOWN. *** Another board shown in the drop-downs closes a
    // table open on the first board's draft - what it holds saved first.
    // ###########################################################################################
    [Fact]
    public async Task Showing_another_board_closes_a_table_open_on_the_first()
    {
        WriteDraft(BoardA);
        WriteDraft(BoardB);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry first = Entry("250407", BoardA);
            HardwareBoardEntry second = Entry("250425", BoardB);
            TabDrafts tab = Tab(first, second);
            tab.ShowBoard(first, Board(), "PAL", null);

            await tab.ToggleTableViewAsync();
            Assert.True(tab.IsTableOpen);

            tab.ShowBoard(second, Board(), "PAL", null);

            for (int attempt = 0; attempt < 50 && tab.IsTableOpen; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.False(tab.IsTableOpen);
            Assert.Equal(BoardB, Assert.IsType<DraftListItem>(tab.GetControl<ContentControl>("CurrentDraftRow").Content).ExcelDataFile);
        });
    }

    // Until something is true or half a second has passed - for what a click starts and does not wait for.
    private static async Task Until(Func<bool> condition)
    {
        for (int attempt = 0; attempt < 50 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }
    }

    // ###########################################################################################
    // Clicking a component opens it in the component pane on the right, in place of the start panel
    // (the plan: "A component opens in a pane, never in a maximized window").
    // ###########################################################################################
    [Fact]
    public async Task Clicking_a_component_opens_it_in_the_pane()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            var window = new Window { Content = tab, Width = 1300, Height = 900 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Button u8 = tab.GetControl<ItemsControl>("ComponentsItemsControl")
                .GetRealizedContainers()
                .Select(container => Avalonia.VisualTree.VisualExtensions.FindDescendantOfType<Button>(container))
                .First(button => button?.Tag is ComponentListItem { BoardLabel: "U8" })!;

            u8.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);

            Assert.True(tab.IsComponentOpen);
            Assert.Equal("U8", tab.ComponentPaneForTests.Label);
            Assert.True(tab.ComponentPaneForTests.IsVisible);
            Assert.False(tab.GetControl<StackPanel>("StartPanel").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // The right side shows one thing: the table opened closes the component pane, a component
    // opened closes the table - and what the pane held is in the draft the table shows.
    // ###########################################################################################
    [Fact]
    public async Task The_table_and_the_component_pane_take_turns_on_the_right()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("FriendlyNameBox").Text = "Processor";

            await tab.ToggleTableViewAsync();
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.IsTableOpen);
            Assert.False(tab.IsComponentOpen);
            Assert.False(tab.ComponentPaneForTests.IsVisible);
            Assert.Equal("Processor", ComponentPaneEdits.RowShown(tab.TableEditor.DraftSessionForTests!.Content.Board, "U8", "PAL")!.FriendlyName);

            await tab.OpenComponentAsync("U8");

            Assert.False(tab.IsTableOpen);
            Assert.True(tab.IsComponentOpen);
            Assert.True(tab.ComponentPaneForTests.IsVisible);
            Assert.False(tab.TableEditor.IsVisible);
        });
    }

    // ###########################################################################################
    // *** CLOSING THE TABLE SETTLES ITS OWN BOARD'S DRAFT, NOT EVERY ONE (code review, 2026-10-26). ***
    // Another board's draft with a change that cannot be saved made closing this table ask about
    // "the table's" changes, every time - and Discard read that other draft's file back, its change
    // gone without it ever being shown. It waits for its own board (or quitting) instead.
    // ###########################################################################################
    [Fact]
    public async Task Closing_the_table_leaves_another_boards_unsaved_draft_alone()
    {
        WriteDraft(BoardA, Board());
        WriteDraft(BoardB, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry, Entry("250425", BoardB));
            bool asked = false;
            tab.UnsavedTableEditsAnswerForTests = _ => { asked = true; return UnsavedTableEditsChoice.Discard; };
            tab.ShowBoard(entry, Board(), "PAL", null);

            // Board B's draft holds a change its file was changed under, so it cannot be saved.
            DraftSession other = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardB)!;
            other.Apply("Waiting", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });
            BoardData elsewhere = Board();
            elsewhere.Components.Add(new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "Added elsewhere" });
            WriteDraft(BoardB, elsewhere);

            await tab.ToggleTableViewAsync();
            Assert.True(tab.IsTableOpen);

            Assert.True(await tab.CloseTableAsync());

            Assert.False(asked);
            Assert.False(tab.IsTableOpen);
            Assert.True(other.HasUnsavedChanges);
            Assert.Equal("906122", ComponentPaneEdits.RowShown(other.Content.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ###########################################################################################
    // *** ANOTHER BOARD SHOWN CLOSES THE TABLE, WHATEVER ITS SAVE DOES (code review, 2026-10-26). ***
    // CRT shows the other board already, so the close cannot be called off: asked, Cancel left the
    // table open on the previous board's draft under the new board's header. What it held stays in
    // that draft, held in memory, for its own board or quitting to save or ask about.
    // ###########################################################################################
    [Fact]
    public async Task Showing_another_board_closes_the_table_without_asking_and_keeps_what_could_not_be_saved()
    {
        WriteDraft(BoardA, Board());
        WriteDraft(BoardB, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entryA = Entry("250407", BoardA);
            HardwareBoardEntry entryB = Entry("250425", BoardB);
            TabDrafts tab = Tab(entryA, entryB);
            bool asked = false;
            tab.UnsavedTableEditsAnswerForTests = _ => { asked = true; return UnsavedTableEditsChoice.Cancel; };
            tab.ShowBoard(entryA, Board(), "PAL", null);

            await tab.ToggleTableViewAsync();
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;
            session.Apply("Typed in the table", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            // Written somewhere else meanwhile, so its save is refused.
            BoardData elsewhere = Board();
            elsewhere.Components.Add(new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "Added elsewhere" });
            WriteDraft(BoardA, elsewhere);

            tab.ShowBoard(entryB, Board(), "PAL", null);
            await Until(() => session.SaveState == DraftSaveState.Failed);

            Assert.False(asked);
            Assert.False(tab.IsTableOpen);
            Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.SaveFailure);
            Assert.Equal("906122", ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ###########################################################################################
    // *** A DRAFT NO EDITOR HOLDS IS SAVED TOO (code review, 2026-10-26). *** Only the table and the
    // pane had a save timer, each for its own draft; a change in a draft neither showed was saved only
    // if whoever let go of it remembered to. The tab's one timer saves every draft with a save due
    // that neither holds - those two save their own, which the board on screen follows.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_no_editor_holds_is_saved_when_due_and_the_panes_is_left_to_the_pane()
    {
        WriteDraft(BoardA, Board());
        WriteDraft(BoardB, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry, Entry("250425", BoardB));
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.AutoSave = false;
            DraftSession shown = tab.ComponentPaneForTests.Session!;
            shown.Apply("In the pane's draft", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            DraftSession other = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardB)!;
            other.Apply("Nobody shows it", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "251715") });

            await tab.SaveDueDraftsNoEditorHoldsAsync(DateTimeOffset.UtcNow.AddMinutes(1));

            Assert.False(other.HasUnsavedChanges);
            Assert.Equal("251715", ComponentPaneEdits.RowShown(BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardB))!.Board, "U8", "PAL")!.PartNumber);
            Assert.True(shown.HasUnsavedChanges);
        });
    }

    // ###########################################################################################
    // Opening the table waits for the pane's close-save before it reads the draft - under CRT's
    // "please wait" (code review, 2026-10-26): the contributor pressed for it, and a large board's
    // save takes a second or two in which the tab took clicks on a half-switched workspace.
    // ###########################################################################################
    [Fact]
    public async Task Opening_the_table_waits_for_the_panes_save_under_please_wait()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            var overlay = new BusyOverlay();
            _ = new Grid { Children = { tab, overlay } };
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");

            var save = new TaskCompletionSource();
            tab.ComponentPaneForTests.UseLastSaveForTests(save.Task);

            Task opening = tab.ToggleTableViewAsync();
            Dispatcher.UIThread.RunJobs();

            Assert.True(overlay.IsBusy);
            Assert.False(tab.IsTableOpen);

            save.SetResult();
            await opening;

            Assert.False(overlay.IsBusy);
            Assert.True(tab.IsTableOpen);
        });
    }

    // ###########################################################################################
    // *** THE PANE IS A VIEW OF THE BOARD SHOWN. *** Another board shown closes it - what it held
    // handed to its draft and saved as it goes.
    // ###########################################################################################
    [Fact]
    public async Task Showing_another_board_closes_the_pane_and_saves_what_it_held()
    {
        WriteDraft(BoardA, Board());
        WriteDraft(BoardB);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry first = Entry("250407", BoardA);
            HardwareBoardEntry second = Entry("250425", BoardB);
            TabDrafts tab = Tab(first, second);
            tab.ShowBoard(first, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            DraftSession session = tab.ComponentPaneForTests.Session!;
            Assert.Equal("U8", tab.ComponentPaneForTests.Label);
            Assert.True(tab.ComponentPaneForTests.GetControl<ScrollViewer>("SectionsScrollViewer").IsVisible);
            tab.ComponentPaneForTests.GetControl<TextBox>("DescriptionBox").Text = "Runs the show";

            tab.ShowBoard(second, Board(), "PAL", null);
            Assert.False(tab.IsComponentOpen);
            Assert.Equal("Runs the show", ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL")!.Description);

            // The save the pane started as it closed - finished before the file is read, and before
            // the next test runs (a save still finishing then refreshed the tab against that test's
            // drafts folder).
            await tab.ComponentPaneForTests.LastSave;
            Dispatcher.UIThread.RunJobs();
            Assert.Null(session.SaveFailure);
            Assert.False(session.HasUnsavedChanges);

            string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA);
            Assert.Equal("Runs the show", ComponentPaneEdits.RowShown(BoardFiles.ReadContent(boardFile)!.Board, "U8", "PAL")!.Description);
        });
    }

    // ###########################################################################################
    // What the pane holds counts as unsaved for quitting and installing an update - and is saved
    // before either (Main.ConfirmLeavingTablesAsync).
    // ###########################################################################################
    [Fact]
    public async Task What_the_pane_holds_counts_as_unsaved_and_is_saved_before_leaving()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.AutoSave = false;
            tab.ComponentPaneForTests.GetControl<TextBox>("PartNumberBox").Text = "906122";

            Assert.True(tab.CommitCellsAndHasUnsavedTableEdits());
            Assert.True(tab.HasUnsavedTableEditsFor(BoardA));
            Assert.True(await tab.SaveDraftEditorsNowAsync());
            Assert.False(tab.CommitCellsAndHasUnsavedTableEdits());

            BoardContent written = BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA))!;
            Assert.Equal("906122", ComponentPaneEdits.RowShown(written.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ###########################################################################################
    // *** A DRAFT NEITHER THE PANE NOR THE TABLE SHOWS STILL COUNTS. *** A save the pane started as
    // it closed can fail (the file held open for a moment); the change then waits in the draft held
    // in memory, and quitting must save it - it asked only about an open pane or table, and the
    // change was gone.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_with_changes_not_saved_that_nothing_shows_is_saved_before_leaving()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardA)!;
            session.Apply("Waiting", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            Assert.False(tab.IsComponentOpen);
            Assert.False(tab.IsTableOpen);
            Assert.True(tab.CommitCellsAndHasUnsavedTableEdits());

            Assert.True(await tab.SaveDraftEditorsNowAsync());
            Assert.False(tab.CommitCellsAndHasUnsavedTableEdits());

            BoardContent written = BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA))!;
            Assert.Equal("906122", ComponentPaneEdits.RowShown(written.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ###########################################################################################
    // *** SUBMIT AND "MANAGE FILES" SETTLE THE DRAFT ITSELF TOO (code review, 2026-10-26). *** They
    // saved only an open table or pane, so a change waiting in the draft held in memory that neither
    // shows - the pane's close-save refused, the header's Undo with nothing open - was left out: the
    // submission sent the board file without it.
    // ###########################################################################################
    [Fact]
    public async Task A_change_nothing_shows_is_saved_before_the_draft_is_submitted()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardA)!;
            session.Apply("Waiting", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            Assert.True(await tab.SettleTableEditsAsync(entry));

            Assert.False(session.HasUnsavedChanges);
            BoardContent written = BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA))!;
            Assert.Equal("906122", ComponentPaneEdits.RowShown(written.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ...and when that save cannot be made, the contributor is asked - and Cancel stops the submit.
    [Fact]
    public async Task A_change_nothing_shows_that_cannot_be_saved_is_asked_about_before_the_draft_is_submitted()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);
            UnsavedTableEditsPrompt? asked = null;
            tab.UnsavedTableEditsAnswerForTests = prompt => { asked = prompt; return UnsavedTableEditsChoice.Cancel; };

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardA)!;
            session.Apply("Waiting", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            // Written somewhere else meanwhile, so the save refuses rather than write over it.
            BoardData elsewhere = Board();
            elsewhere.Components.Add(new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "Added elsewhere" });
            WriteDraft(BoardA, elsewhere);

            Assert.False(await tab.SettleTableEditsAsync(entry));

            Assert.Equal(UnsavedTableEditsPrompt.DraftChangedOnDisk, asked);
            Assert.True(session.HasUnsavedChanges);
        });
    }

    // ###########################################################################################
    // The left column follows the pane at once - a name typed there is in the list, with the draft's
    // "Changed" mark - while the board on screen catches up later.
    // ###########################################################################################
    [Fact]
    public async Task The_left_column_follows_the_pane()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("FriendlyNameBox").Text = "Processor";
            tab.ComponentPaneForTests.CommitHeld();

            ComponentListItem u8 = tab.ComponentsShownForTests.Single(item => item.BoardLabel == "U8");
            Assert.Contains("Processor", u8.DisplayText, StringComparison.Ordinal);
            Assert.True(u8.IsDrafted);
        });
    }

    // ###########################################################################################
    // A draft discarded while its component is open: the pane lets go - nothing handed over, nothing
    // written back into the folder that is gone.
    // ###########################################################################################
    [Fact]
    public async Task Discarding_the_draft_lets_go_of_the_pane_without_writing_it_back()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("FriendlyNameBox").Text = "Processor";

            tab.DiscardConfirmed(entry);

            Assert.False(tab.IsComponentOpen);
            Assert.True(tab.ComponentPaneForTests.LastSave.IsCompleted);
            Assert.False(File.Exists(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA)));
        });
    }

    // ###########################################################################################
    // Why no draft was made is about the board it was pressed on: showing another board clears it
    // (code review, 2026-10-09); reading the same board again does not.
    // ###########################################################################################
    [Fact]
    public void A_problem_making_the_draft_is_cleared_when_another_board_is_shown()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();
            HardwareBoardEntry first = Entry("250407", BoardA);
            tab.ShowBoard(first, Board(), "PAL", null);

            tab.ShowDraftProblem("No draft could be made of this board: the published board could not be read.");
            tab.ShowBoard(first, Board(), "PAL", null);
            Assert.StartsWith("No draft could be made", tab.DraftProblemForTests);

            tab.ShowBoard(Entry("250425", BoardB), Board(), "PAL", null);
            Assert.Null(tab.DraftProblemForTests);
        });
    }
}
