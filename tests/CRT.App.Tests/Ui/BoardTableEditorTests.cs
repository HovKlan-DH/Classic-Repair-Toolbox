using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Drafts tab's table editor as a control (owner request, 2026-09-24) - what the grid
// actually shows. The rules behind it (pairing, colours, ghost placement, what a save writes)
// are pinned in CRT.Data.Tests' BoardTableDocumentTests; this file pins that the control PAINTS
// them: that a changed cell's container really carries the orange wash, a deleted row really
// carries the strike-through class, and the toolbar and clipboard act on the right cell.
//
// *** THE COLOUR TESTS REALISE THE GRID IN A SHOWN WINDOW. *** The colour is a binding inside a
// per-column cell theme, so the only proof it works is a real DataGridCell with a real
// Background - a test reading the model's State would pass against a grid that painted nothing.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableEditorTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private const string SystemKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string WorkbookPath => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardTableEditorTests.SystemKey);

    public void Dispose() => this.thisWorkspace.Dispose();

    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, TechnicalNameOrValue = "x" };

    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    private void WriteDraft(BoardData board)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(this.WorkbookPath)!);
        BoardWorkbookWriter.Write(this.WorkbookPath, board);
    }

    // The editor loaded on the draft, with Components selected (the first sheet is schematics).
    private BoardTableEditor OpenEditor(BoardData draft, BoardData? published)
    {
        this.WriteDraft(draft);

        var editor = new BoardTableEditor();
        Assert.True(editor.Load(this.DraftsRoot, BoardTableEditorTests.SystemKey, published));
        editor.SelectSheet(Components(editor));

        return editor;
    }

    private static BoardTableSheet Components(BoardTableEditor editor) =>
        editor.SessionForTests!.Document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

    private static int Column(string name) =>
        BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(name);

    private static BoardTableRow Row(BoardTableEditor editor, string label, bool deleted = false) =>
        Components(editor).Rows.Single(row =>
            row.IsDeleted == deleted && row.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text == label);

    private static Window Show(BoardTableEditor editor)
    {
        var window = new Window { Content = editor, Width = 1200, Height = 600 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    private static Color? ThemeColor(string key)
    {
        Application app = Application.Current!;
        return app.TryGetResource(key, app.ActualThemeVariant, out object? resource) && resource is ISolidColorBrush brush
            ? brush.Color
            : null;
    }

    private static DataGridCell CellOnScreen(Window window, BoardTableRow row, int columnIndex) =>
        window.GetVisualDescendants()
            .OfType<DataGridCell>()
            .Single(cell => ReferenceEquals(cell.DataContext, row) && cell.OwningColumn?.Tag is int tag && tag == columnIndex);

    // ------------------------------------------------------------------ Shape

    private static TabControl SheetTabs(BoardTableEditor editor) => editor.GetControl<TabControl>("SheetTabs");

    private static TabItem SheetTab(BoardTableEditor editor, string header) =>
        SheetTabs(editor).Items.OfType<TabItem>().Single(tab => (tab.Header as string) == header);

    [Fact]
    public void Every_workbook_sheet_gets_a_real_TAB_and_a_changed_sheet_names_its_count()
    {
        // Real TabItems in a TabControl (owner request, 2026-09-24 - they were buttons at
        // first), so they look and behave like the application's own tabs.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "CPU 6510"), Component("U2")),
                published: Board(Component("U1", "CPU")));

            var tabs = SheetTabs(editor).Items.OfType<TabItem>().ToList();

            Assert.Equal(BoardWorkbookSchema.AllSheets.Count, tabs.Count);
            Assert.Contains(tabs, tab => (tab.Header as string) == "Components (2)");
            Assert.Contains(tabs, tab => (tab.Header as string) == "Credits");

            // The shown sheet's tab is the selected one, and the tabs carry no content of their
            // own - the one grid below shows the selected sheet.
            Assert.Same(SheetTab(editor, "Components (2)"), SheetTabs(editor).SelectedItem);
            Assert.All(tabs, tab => Assert.Null(tab.Content));
        });
    }

    [Fact]
    public void A_sheet_tab_label_follows_its_change_count_as_the_table_is_edited()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: Board(Component("U1", "CPU")));

            Assert.NotNull(SheetTab(editor, "Components"));

            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "CPU 6510";
            editor.RefreshPendingForTests();

            Assert.NotNull(SheetTab(editor, "Components (1)"));
        });
    }

    [Fact]
    public void The_grid_shows_a_marker_column_then_the_sheets_own_columns_in_order()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            Assert.Equal(
                BoardWorkbookSchema.Components.ColumnOrder,
                grid.Columns.Skip(1).Select(column => column.Header as string));

            Assert.True(grid.Columns[0].IsReadOnly);

            // Through a view (so "Show changes only" can filter it), over the sheet's own rows.
            Assert.Same(Components(editor).Rows, Assert.IsType<Avalonia.Collections.DataGridCollectionView>(grid.ItemsSource).SourceCollection);
        });
    }

    [Fact]
    public void A_single_click_only_selects_a_cell_so_select_then_paste_works_as_in_Excel()
    {
        // The grid's own default edits on a single click, which would open the cell's editor
        // before Ctrl+V on a "selected" cell could ever happen.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            DataGridEditTriggers triggers = editor.GetControl<DataGrid>("TableGrid").EditTriggers;

            Assert.False(triggers.HasFlag(DataGridEditTriggers.CellClick));
            Assert.True(triggers.HasFlag(DataGridEditTriggers.CellDoubleClick));
            Assert.True(triggers.HasFlag(DataGridEditTriggers.TextInput));
            Assert.True(triggers.HasFlag(DataGridEditTriggers.F2));
        });
    }

    [Fact]
    public void Picking_another_sheet_tab_shows_that_sheet()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);

            SheetTabs(editor).SelectedItem = SheetTab(editor, BoardWorkbookSchema.SheetCredits);

            Assert.Equal(BoardWorkbookSchema.SheetCredits, editor.CurrentSheet!.Name);
            Assert.Equal(
                BoardWorkbookSchema.Credits.ColumnOrder,
                editor.GetControl<DataGrid>("TableGrid").Columns.Skip(1).Select(column => column.Header as string));
        });
    }

    // ------------------------------------------------------------------ Colours on screen

    [Fact]
    public void A_changed_cell_is_painted_orange_and_its_unchanged_neighbour_is_not()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "CPU 6510")),
                published: Board(Component("U1", "CPU")));
            Window window = Show(editor);

            BoardTableRow row = Row(editor, "U1");

            DataGridCell changed = CellOnScreen(window, row, Column(BoardWorkbookSchema.ColFriendlyName));
            DataGridCell unchanged = CellOnScreen(window, row, Column(BoardWorkbookSchema.ColBoardLabel));

            Assert.Equal(ThemeColor("BoardTable_Modified_Bg"), (changed.Background as ISolidColorBrush)?.Color);
            Assert.NotEqual(ThemeColor("BoardTable_Modified_Bg"), (unchanged.Background as ISolidColorBrush)?.Color);
            Assert.Equal("Published value: CPU", ToolTip.GetTip(changed));

            window.Close();
        });
    }

    [Fact]
    public void An_added_row_is_green_and_a_deleted_row_is_red_and_struck_through()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1"), Component("U9")),
                published: Board(Component("U1"), Component("U2")));
            Window window = Show(editor);

            BoardTableRow added = Row(editor, "U9");
            BoardTableRow ghost = Row(editor, "U2", deleted: true);

            Assert.Equal(
                ThemeColor("BoardTable_Added_Bg"),
                (CellOnScreen(window, added, 0).Background as ISolidColorBrush)?.Color);
            Assert.Equal(
                ThemeColor("BoardTable_Deleted_Bg"),
                (CellOnScreen(window, ghost, 0).Background as ISolidColorBrush)?.Color);

            DataGridRow ghostRow = window.GetVisualDescendants().OfType<DataGridRow>().Single(r => ReferenceEquals(r.DataContext, ghost));
            Assert.Contains("BoardTableDeleted", ghostRow.Classes);

            // The class alone proves nothing: the grid draws cell text with DataGridSearchTextBlock,
            // a TextBlock SUBCLASS, and a plain "TextBlock" style selector matches only the exact
            // type - so the strike-through once matched nothing while the class sat there looking
            // right. Read the decoration off the text actually drawn.
            TextBlock ghostText = CellOnScreen(window, ghost, Column(BoardWorkbookSchema.ColBoardLabel))
                .GetVisualDescendants().OfType<TextBlock>().Single();
            Assert.Contains(ghostText.TextDecorations ?? [], d => d.Location == TextDecorationLocation.Strikethrough);

            TextBlock addedText = CellOnScreen(window, added, Column(BoardWorkbookSchema.ColBoardLabel))
                .GetVisualDescendants().OfType<TextBlock>().Single();
            Assert.DoesNotContain(addedText.TextDecorations ?? [], d => d.Location == TextDecorationLocation.Strikethrough);

            DataGridRow addedRow = window.GetVisualDescendants().OfType<DataGridRow>().Single(r => ReferenceEquals(r.DataContext, added));
            Assert.DoesNotContain("BoardTableDeleted", addedRow.Classes);

            window.Close();
        });
    }

    [Fact]
    public void A_cell_edit_recolours_the_cell_once_the_scheduled_refresh_has_run()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: Board(Component("U1", "CPU")));
            Window window = Show(editor);

            BoardTableRow row = Row(editor, "U1");
            row.Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "CPU 6510";

            // The refresh is posted to the dispatcher rather than run inside the edit.
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                ThemeColor("BoardTable_Modified_Bg"),
                (CellOnScreen(window, row, Column(BoardWorkbookSchema.ColFriendlyName)).Background as ISolidColorBrush)?.Color);
            Assert.True(editor.GetControl<Button>("SaveButton").IsEnabled);

            window.Close();
        });
    }

    [Fact]
    public void With_nothing_published_only_the_flagged_pill_is_shown_and_the_table_says_why()
    {
        // Nothing to be added, changed or deleted AGAINST - but a duplicate is a duplicate either
        // way, so the flagged pill stays and still counts.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U1")), published: null);

            Assert.False(editor.GetControl<Border>("AddedPill").IsVisible);
            Assert.False(editor.GetControl<Border>("ModifiedPill").IsVisible);
            Assert.False(editor.GetControl<Border>("DeletedPill").IsVisible);
            Assert.False(editor.GetControl<CheckBox>("OnlyChangesCheckBox").IsVisible);
            Assert.True(editor.GetControl<Border>("FlaggedPill").IsVisible);
            Assert.Equal("1", editor.GetControl<TextBlock>("FlaggedCountText").Text);

            TextBlock status = editor.GetControl<TextBlock>("StatusText");
            Assert.True(status.IsVisible);
            Assert.Contains("no row is marked as added, changed or deleted", status.Text);
        });
    }

    [Fact]
    public void With_a_published_board_all_four_pills_and_the_filter_are_shown()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: Board(Component("U1")));

            Assert.All(
                ["AddedPill", "ModifiedPill", "DeletedPill", "FlaggedPill"],
                name => Assert.True(editor.GetControl<Border>(name).IsVisible, name));
            Assert.True(editor.GetControl<CheckBox>("OnlyChangesCheckBox").IsVisible);
        });
    }

    // ------------------------------------------------------------------ Keyboard

    [Fact]
    public void Focusing_the_table_puts_the_keyboard_on_its_first_cell()
    {
        // What the Drafts tab does when a table opens or the tab is shown again, so typing goes
        // into the table instead of the always-on component filter on the left.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = Show(editor);

            editor.FocusGrid();
            Dispatcher.UIThread.RunJobs();

            Assert.True(editor.GetControl<DataGrid>("TableGrid").IsKeyboardFocusWithin);
            Assert.Same(Row(editor, "U1"), editor.CurrentRow);

            window.Close();
        });
    }

    [Fact]
    public void Typing_on_a_selected_EMPTY_cell_puts_the_text_in_that_cell()
    {
        // The project owner's own steps (2026-09-24): select an empty Friendly name cell and type.
        // Driven with REAL key input through the headless platform, so it exercises the grid's
        // "typing starts editing" trigger rather than setting the model behind its back.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("C1")), published: null);
            Window window = Show(editor);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            editor.SelectCell(Row(editor, "C1"), Column(BoardWorkbookSchema.ColFriendlyName));
            grid.Focus();
            Dispatcher.UIThread.RunJobs();

            window.KeyTextInput("Capacitor");
            Dispatcher.UIThread.RunJobs();
            grid.CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Capacitor", Row(editor, "C1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);
            Assert.True(editor.HasUnsavedChanges);

            window.Close();
        });
    }

    [Fact]
    public void Every_data_column_is_editable_and_only_the_marker_column_is_not()
    {
        // *** THE REPORTED BUG (2026-09-24): nothing could be typed into any cell. *** Left to
        // itself the grid decides read-only-ness from each column's binding path, and "Cells[i]"
        // indexes a read-only list, so EVERY data column came out read-only. Every other test in
        // this file set cell text through the model, which is why none of them noticed.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            Assert.True(grid.Columns[0].IsReadOnly);
            Assert.All(grid.Columns.Skip(1), column => Assert.False(column.IsReadOnly, $"[{column.Header}] is read-only"));
        });
    }

    [Fact]
    public void F2_opens_the_selected_cell_for_editing_but_never_a_deleted_rows_cell()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: Board(Component("U1"), Component("U2")));
            Window window = Show(editor);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            grid.Focus();
            window.KeyPress(Avalonia.Input.Key.F2, Avalonia.Input.RawInputModifiers.None, Avalonia.Input.PhysicalKey.F2, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.Contains(grid.GetVisualDescendants().OfType<TextBox>(), textBox => textBox.IsFocused);
            grid.CancelEdit();
            Dispatcher.UIThread.RunJobs();

            // A ghost shows what was published; OnBeginningEdit refuses to open it.
            editor.SelectCell(Row(editor, "U2", deleted: true), Column(BoardWorkbookSchema.ColFriendlyName));
            grid.Focus();
            Assert.False(grid.BeginEdit());

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Toolbar

    [Fact]
    public void Insert_row_below_adds_an_empty_row_straight_after_the_selected_one_and_selects_it()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            editor.InsertRowBelow();

            BoardTableRow inserted = Components(editor).Rows[1];
            Assert.True(inserted.IsBlank);
            Assert.Same(inserted, editor.CurrentRow);
            Assert.True(editor.HasUnsavedChanges);

            window.Close();
        });
    }

    [Fact]
    public void Insert_row_above_adds_an_empty_row_straight_before_the_selected_one_and_selects_it()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U2"), Column(BoardWorkbookSchema.ColFriendlyName));
            editor.InsertRowAbove();

            BoardTableRow inserted = Components(editor).Rows[1];
            Assert.True(inserted.IsBlank);
            Assert.Same(inserted, editor.CurrentRow);
            Assert.Same(Row(editor, "U2"), Components(editor).Rows[2]);

            window.Close();
        });
    }

    [Fact]
    public void The_insert_buttons_say_above_and_below_and_there_are_no_move_buttons()
    {
        // Owner request, 2026-09-24: "Insert row" became two buttons, and Move up / Move down
        // went - a row is dragged by its grip, or moved with Alt+Up / Alt+Down.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);

            Assert.Equal("Insert row above", editor.GetControl<Button>("InsertRowAboveButton").Content);
            Assert.Equal("Insert row below", editor.GetControl<Button>("InsertRowBelowButton").Content);
            Assert.Null(editor.FindControl<Button>("InsertRowButton"));
            Assert.Null(editor.FindControl<Button>("MoveRowUpButton"));
            Assert.Null(editor.FindControl<Button>("MoveRowDownButton"));
        });
    }

    [Fact]
    public void The_toolbar_only_offers_what_applies_to_the_selected_cell()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "CPU 6510"), Component("U3")),
                published: Board(Component("U1", "CPU"), Component("U2")));
            Window window = Show(editor);

            Button delete = editor.GetControl<Button>("DeleteRowButton");

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            Assert.True(delete.IsEnabled);

            // A red row is already deleted.
            editor.SelectCell(Row(editor, "U2", deleted: true), 0);
            Assert.False(delete.IsEnabled);

            window.Close();
        });
    }

    [Theory]
    [InlineData("RestoreRowButton")]
    [InlineData("RevertCellButton")]
    public void There_are_no_Restore_row_or_Revert_cell_buttons_undo_does_both(string removed)
    {
        // Removed at the owner's request (2026-09-24) once Ctrl+Z existed: less clutter. The
        // model keeps BoardTableSheet.RestoreRow / RevertCell for the maintainer application's table.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: Board(Component("U1")));

            Assert.Null(editor.FindControl<Button>(removed));
        });
    }

    [Fact]
    public void Delete_then_Ctrl_Z_puts_a_published_row_back_where_it_was()
    {
        UiTest.Run(() =>
        {
            BoardData board = Board(Component("U1"), Component("U2"), Component("U3"));
            BoardTableEditor editor = this.OpenEditor(board, published: board);
            Window window = Show(editor);
            BoardTableRow u2 = Row(editor, "U2");

            editor.SelectCell(u2, 0);
            editor.DeleteRow();
            Assert.True(Components(editor).Rows[1].IsDeleted);

            editor.Undo();

            Assert.Same(u2, Components(editor).Rows[1]);
            Assert.Equal(0, Components(editor).ChangeCount);

            window.Close();
        });
    }

    // ###########################################################################################
    // A deleted component takes its rows on the other sheets with it (2026-09-25), and those sheets
    // are not the one on screen - so the status line says what else went, and that undo brings it
    // back.
    // ###########################################################################################
    [Fact]
    public void Deleting_a_component_says_what_went_with_it_on_the_other_sheets()
    {
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                Components = [Component("U1"), Component("U2")],
                ComponentImages = [new ComponentImageEntry { BoardLabel = "U1", Name = "Clock", File = "Commodore/C64/250407/u1.png" }]
            };

            BoardTableEditor editor = this.OpenEditor(board, published: board);
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), 0);
            editor.DeleteRow();

            Assert.Equal(
                "Also deleted with U1: 1 row on Component images. Undo (Ctrl+Z) brings it all back.",
                editor.GetControl<TextBlock>("StatusText").Text);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Clipboard

    [Fact]
    public void Pasting_one_Excel_cell_replaces_the_selected_cell_without_Excels_line_break()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));

            Assert.True(editor.PasteIntoCurrentCell("CPU 6510\r\n"));
            Assert.Equal("CPU 6510", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);

            window.Close();
        });
    }

    [Fact]
    public void Pasting_a_block_of_cells_changes_nothing_and_says_why()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: null);
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));

            Assert.False(editor.PasteIntoCurrentCell("U8\tCPU\r\nU9\tVIC\r\n"));
            Assert.Equal("CPU", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);
            Assert.Contains("one cell at a time", editor.GetControl<TextBlock>("StatusText").Text);

            window.Close();
        });
    }

    [Fact]
    public void Pasting_onto_a_deleted_row_is_refused()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(), published: Board(Component("U2", "VIC")));
            Window window = Show(editor);

            BoardTableRow ghost = Row(editor, "U2", deleted: true);
            editor.SelectCell(ghost, Column(BoardWorkbookSchema.ColFriendlyName));

            Assert.False(editor.PasteIntoCurrentCell("Typed over"));
            Assert.Equal("VIC", ghost.Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);

            window.Close();
        });
    }

    [Fact]
    public void Copying_a_cell_gives_its_text_quoted_only_when_Excel_needs_it()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "5.25\" drive")), published: null);
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColBoardLabel));
            Assert.Equal("U1", editor.CopyTextOfCurrentCell());

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            Assert.Equal("\"5.25\"\" drive\"", editor.CopyTextOfCurrentCell());

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Saving

    [Fact]
    public void Save_writes_the_draft_says_so_and_raises_Saved()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: null);
            bool savedRaised = false;
            editor.Saved += (_, _) => savedRaised = true;

            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "CPU 6510";

            Assert.Equal(DraftWorkbookEditOutcome.Saved, editor.Save());

            Assert.True(savedRaised);
            Assert.False(editor.HasUnsavedChanges);
            Assert.Equal("Saved.", editor.GetControl<TextBlock>("StatusText").Text);
            Assert.Equal(
                "CPU 6510",
                DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, BoardTableEditorTests.SystemKey)!.Components.Single().FriendlyName);
        });
    }

    [Fact]
    public void Save_is_refused_with_an_explanation_when_the_draft_changed_outside_the_table()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: null);
            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "From the table";

            // Edited in Excel meanwhile.
            this.WriteDraft(Board(Component("U1", "From Excel")));

            Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, editor.Save());

            Assert.Contains("changed outside this table", editor.GetControl<TextBlock>("StatusText").Text);
            Assert.True(editor.HasUnsavedChanges);
            Assert.Equal(
                "From Excel",
                DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, BoardTableEditorTests.SystemKey)!.Components.Single().FriendlyName);
        });
    }

    [Fact]
    public void Catching_up_reloads_silently_when_nothing_is_unsaved_but_only_warns_when_something_is()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: null);

            this.WriteDraft(Board(Component("U1", "From Excel")));
            editor.CheckDraftFile();

            Assert.Equal("From Excel", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);
            Assert.Contains("Updated", editor.GetControl<TextBlock>("StatusText").Text);
            Assert.False(editor.GetControl<Border>("ChangedOnDiskBar").IsVisible);

            // Now with an unsaved edit: reloading would lose it, so the table only warns - with the
            // warning BAR, and Save greyed out, since that save would be refused.
            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "Unsaved";
            editor.RefreshPendingForTests();
            this.WriteDraft(Board(Component("U1", "Excel again")));
            editor.CheckDraftFile();

            Assert.Equal("Unsaved", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);
            Assert.True(editor.GetControl<Border>("ChangedOnDiskBar").IsVisible);
            Assert.Contains("out of date", editor.GetControl<TextBlock>("ChangedOnDiskText").Text);
            Assert.False(editor.GetControl<Button>("SaveButton").IsEnabled);

            // Reload (with the edits given up) takes the bar away and shows the file.
            editor.Reload();

            Assert.Equal("Excel again", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);
            Assert.False(editor.GetControl<Border>("ChangedOnDiskBar").IsVisible);
        });
    }

    // ------------------------------------------------------------------ Palette

    [Theory]
    [InlineData("BoardTable_Added_Bg")]
    [InlineData("BoardTable_Modified_Bg")]
    [InlineData("BoardTable_Deleted_Bg")]
    [InlineData("BoardTable_Flagged_Bg")]
    [InlineData("BoardTable_CurrentCell_Border")]
    [InlineData("BoardTable_Notice_Bg")]
    [InlineData("BoardTable_Notice_Border")]
    public void Every_table_colour_is_defined_in_BOTH_themes(string key)
    {
        // A key missing from one theme renders the converter's hardcoded fallback there, which
        // nothing reports - see WorkbooksPaletteTests for the same trap.
        UiTest.Run(() =>
        {
            Application app = Application.Current!;

            Assert.True(app.TryGetResource(key, ThemeVariant.Light, out object? light) && light is ISolidColorBrush, $"{key} missing from Light");
            Assert.True(app.TryGetResource(key, ThemeVariant.Dark, out object? dark) && dark is ISolidColorBrush, $"{key} missing from Dark");
        });
    }

    // ------------------------------------------------------------------ Moving rows (2026-09-24)

    private static List<string> Labels(BoardTableEditor editor) =>
        Components(editor).Rows.Select(row => row.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text).ToList();

    [Fact]
    public void Alt_Up_and_Alt_Down_move_the_selected_row_and_the_cursor_goes_with_it()
    {
        // The buttons are gone (2026-09-24); the keys that did the same remain.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2"), Component("U3")), published: null);
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));

            editor.MoveRowDown();

            Assert.Equal(["U2", "U1", "U3"], Labels(editor));
            Assert.Same(Row(editor, "U1"), editor.CurrentRow);
            Assert.Equal(Column(BoardWorkbookSchema.ColFriendlyName), editor.CurrentCell!.ColumnIndex);
            Assert.True(editor.HasUnsavedChanges);

            editor.MoveRowUp();
            Assert.Equal(["U1", "U2", "U3"], Labels(editor));

            window.Close();
        });
    }

    [Fact]
    public void Alt_Down_moves_the_selected_row_with_a_real_keypress()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = Show(editor);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            editor.SelectCell(Row(editor, "U1"), 0);
            grid.Focus();
            window.KeyPress(Avalonia.Input.Key.Down, Avalonia.Input.RawInputModifiers.Alt, Avalonia.Input.PhysicalKey.ArrowDown, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["U2", "U1"], Labels(editor));

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Dragging a row (2026-09-24)

    private static Avalonia.Controls.Primitives.DataGridRowHeader HeaderOf(Window window, BoardTableRow row) =>
        window.GetVisualDescendants().OfType<Avalonia.Controls.Primitives.DataGridRowHeader>()
            .Single(header => ReferenceEquals(header.DataContext, row));

    private static DataGridRow ContainerOf(Window window, BoardTableRow row) =>
        window.GetVisualDescendants().OfType<DataGridRow>().Single(container => ReferenceEquals(container.DataContext, row));

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    private static void PressAt(Window window, Point point)
    {
        window.MouseDown(point, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    private static void DragTo(Window window, Point point)
    {
        window.MouseMove(point, Avalonia.Input.RawInputModifiers.LeftMouseButton);
        Dispatcher.UIThread.RunJobs();
    }

    private static void ReleaseAt(Window window, Point point)
    {
        window.MouseUp(point, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    [Fact]
    public void Dragging_a_grip_turns_the_row_into_a_dashed_placeholder_that_travels_and_drops_there()
    {
        // Owner request: "a placeholder when I drag a row, like moving an image in the
        // worklog". Driven with real pointer input: press the grip, move onto another row, release.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1"), Component("U2"), Component("U3"), Component("U4")),
                published: null);
            Window window = Show(editor);
            BoardTableRow u1 = Row(editor, "U1");

            Point grip = CentreOf(window, HeaderOf(window, u1));
            Point overU3 = CentreOf(window, CellOnScreen(window, Row(editor, "U3"), Column(BoardWorkbookSchema.ColFriendlyName)));

            PressAt(window, grip);
            DragTo(window, overU3);

            // Live: U1 already sits where U3 was, the others shifted up into the order a drop gives.
            Assert.Same(u1, editor.DraggedRowForTests);
            Assert.Equal(["U2", "U3", "U1", "U4"], Labels(editor));

            // ... drawn as the worklog's placeholder: the row's own background as a red dashed slot,
            // its cells faded out so only the slot shows.
            DataGridRow container = ContainerOf(window, u1);
            Assert.Contains("BoardTableDropPlaceholder", container.Classes);

            Avalonia.Controls.Shapes.Rectangle slot = TemplatePart<Avalonia.Controls.Shapes.Rectangle>(container, "BackgroundRectangle");
            Assert.Equal(2, slot.StrokeThickness);
            Assert.NotEmpty(slot.StrokeDashArray!);
            Assert.Equal(ThemeColor("Text_Fail_Fg"), (slot.Stroke as ISolidColorBrush)?.Color);
            Assert.All(container.GetVisualDescendants().OfType<DataGridCell>(), cell => Assert.Equal(0, cell.Opacity));

            // No other row is a placeholder.
            Assert.Single(window.GetVisualDescendants().OfType<DataGridRow>(), row => row.Classes.Contains("BoardTableDropPlaceholder"));

            ReleaseAt(window, overU3);

            Assert.Null(editor.DraggedRowForTests);
            Assert.Equal(["U2", "U3", "U1", "U4"], Labels(editor));
            Assert.DoesNotContain(window.GetVisualDescendants().OfType<DataGridRow>(), row => row.Classes.Contains("BoardTableDropPlaceholder"));
            Assert.Same(u1, editor.CurrentRow);

            // One gesture, one undo step.
            Assert.Equal(1, editor.SessionForTests!.Document.History.UndoCount);
            editor.Undo();
            Assert.Equal(["U1", "U2", "U3", "U4"], Labels(editor));

            window.Close();
        });
    }

    [Fact]
    public void A_press_on_the_grip_without_moving_is_just_a_click()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = Show(editor);
            Point grip = CentreOf(window, HeaderOf(window, Row(editor, "U2")));

            PressAt(window, grip);
            DragTo(window, grip + new Point(1, 2));
            ReleaseAt(window, grip);

            Assert.Equal(["U1", "U2"], Labels(editor));
            Assert.Equal(0, editor.SessionForTests!.Document.History.UndoCount);
            Assert.False(editor.HasUnsavedChanges);
            Assert.Same(Row(editor, "U2"), editor.CurrentRow);

            window.Close();
        });
    }

    [Fact]
    public void Dragging_above_the_rows_carries_the_row_up_one_step_per_move()
    {
        // Past the edge of the rows on screen there is no row to move onto; each move there takes
        // it one live row further, so it can be carried beyond what is visible.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2"), Component("U3")), published: null);
            Window window = Show(editor);
            BoardTableRow u3 = Row(editor, "U3");
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            Point grip = CentreOf(window, HeaderOf(window, u3));
            Point aboveRows = grid.TranslatePoint(new Point(grip.X, 4), window)!.Value;

            PressAt(window, grip);
            DragTo(window, aboveRows);
            Assert.Equal(["U1", "U3", "U2"], Labels(editor));

            DragTo(window, aboveRows + new Point(0, 1));
            Assert.Equal(["U3", "U1", "U2"], Labels(editor));

            ReleaseAt(window, aboveRows);
            Assert.Equal(1, editor.SessionForTests!.Document.History.UndoCount);

            window.Close();
        });
    }

    [Fact]
    public void A_red_deleted_row_has_no_grip_to_drag_and_is_not_a_place_to_drop_onto()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1"), Component("U3")),
                published: Board(Component("U1"), Component("U2"), Component("U3")));
            Window window = Show(editor);
            BoardTableRow ghost = Row(editor, "U2", deleted: true);

            // Pressing a ghost's header starts nothing.
            Point ghostHeader = CentreOf(window, HeaderOf(window, ghost));
            PressAt(window, ghostHeader);
            DragTo(window, CentreOf(window, CellOnScreen(window, Row(editor, "U3"), 0)));
            ReleaseAt(window, ghostHeader);
            Assert.Equal(0, editor.SessionForTests!.Document.History.UndoCount);

            // Dragging U3 over the ghost leaves everything where it is.
            PressAt(window, CentreOf(window, HeaderOf(window, Row(editor, "U3"))));
            Point overGhost = CentreOf(window, CellOnScreen(window, ghost, 0));
            DragTo(window, overGhost);
            Assert.Equal(["U1", "U2", "U3"], Labels(editor));
            Assert.True(Components(editor).Rows[1].IsDeleted);
            ReleaseAt(window, overGhost);

            window.Close();
        });
    }

    [Fact]
    public void Nothing_can_be_dragged_while_show_changes_only_hides_rows()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "one"), Component("U2", "two")),
                published: Board(Component("U1"), Component("U2")));
            Window window = Show(editor);
            editor.OnlyChanges = true;
            Dispatcher.UIThread.RunJobs();

            Point grip = CentreOf(window, HeaderOf(window, Row(editor, "U2")));
            PressAt(window, grip);
            DragTo(window, CentreOf(window, CellOnScreen(window, Row(editor, "U1"), 0)));

            Assert.Null(editor.DraggedRowForTests);
            Assert.Equal(["U1", "U2"], Labels(editor));

            ReleaseAt(window, grip);
            window.Close();
        });
    }

    // ------------------------------------------------------------------ Save button and message

    [Fact]
    public void Save_changes_wears_the_same_red_as_the_Discard_button()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Window window = Show(editor);
            Button save = editor.GetControl<Button>("SaveButton");

            Assert.Equal(ThemeColor("Button_Cancel_Bg"), (save.Background as ISolidColorBrush)?.Color);
            Assert.Equal(ThemeColor("Button_Cancel_Fg"), (save.Foreground as ISolidColorBrush)?.Color);
            Assert.Equal(ThemeColor("Button_Cancel_Border"), (save.BorderBrush as ISolidColorBrush)?.Color);

            window.Close();
        });
    }

    [Fact]
    public void Saving_a_new_component_says_it_was_put_into_its_category()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                new BoardData
                {
                    Components =
                    [
                        new ComponentEntry { BoardLabel = "U1", TechnicalNameOrValue = "x", Category = "IC" },
                        new ComponentEntry { BoardLabel = "U3", TechnicalNameOrValue = "x", Category = "IC" },
                    ],
                },
                published: null);

            BoardTableRow inserted = Components(editor).InsertRow(null);
            inserted.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text = "U2";
            inserted.Cells[Column(BoardWorkbookSchema.ColCategory)].Text = "IC";
            editor.RefreshPendingForTests();

            Assert.Equal(DraftWorkbookEditOutcome.Saved, editor.Save());

            Assert.Contains("New components were put into their category", editor.GetControl<TextBlock>("StatusText").Text);
            Assert.Equal(["U1", "U2", "U3"], Labels(editor));
        });
    }

    [Fact]
    public void Every_real_row_shows_its_drag_grip_and_a_deleted_row_shows_none()
    {
        // The grid's template hides the grip (it shows it only while hovering, and only with its own
        // row drag on, which the editor leaves off). The editor's styles show it always - and they
        // need a class trigger to beat the template's hidden value, so a style that silently loses
        // is exactly what this catches.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U2")), published: Board(Component("U1"), Component("U2")));
            Window window = Show(editor);

            Border GripOf(BoardTableRow row) =>
                window.GetVisualDescendants()
                    .OfType<Avalonia.Controls.Primitives.DataGridRowHeader>()
                    .Single(header => ReferenceEquals(header.DataContext, row))
                    .GetVisualDescendants().OfType<Border>().Single(border => border.Name == "DragGrip");

            Assert.True(GripOf(Row(editor, "U2")).IsVisible);
            Assert.False(GripOf(Row(editor, "U1", deleted: true)).IsVisible);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Round 3 (2026-09-24)

    // A window with a control BEFORE the table, as the main window has the hardware drop-down -
    // so a Tab that escaped the grid would visibly land somewhere else.
    private static Window ShowWithControlBefore(BoardTableEditor editor)
    {
        var window = new Window
        {
            Content = new DockPanel { Children = { new ComboBox { Width = 100 }, editor } },
            Width = 1300,
            Height = 600,
        };

        window.Show();
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    private static void Press(Window window, Avalonia.Input.Key key, Avalonia.Input.RawInputModifiers modifiers = Avalonia.Input.RawInputModifiers.None)
    {
        window.KeyPress(key, modifiers, key == Avalonia.Input.Key.Tab ? Avalonia.Input.PhysicalKey.Tab : Avalonia.Input.PhysicalKey.None, keySymbol: null);
        Dispatcher.UIThread.RunJobs();
    }

    [Fact]
    public void Tab_moves_one_cell_right_and_Shift_Tab_one_cell_left_staying_in_the_table()
    {
        // Reported: Tab took the cursor out of the table to the hardware drop-down.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = ShowWithControlBefore(editor);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            editor.SelectCell(Row(editor, "U1"), 0);
            grid.Focus();

            Press(window, Avalonia.Input.Key.Tab);
            Assert.Equal(1, editor.CurrentCell!.ColumnIndex);
            Assert.True(grid.IsKeyboardFocusWithin);

            Press(window, Avalonia.Input.Key.Tab, Avalonia.Input.RawInputModifiers.Shift);
            Assert.Equal(0, editor.CurrentCell!.ColumnIndex);
            Assert.True(grid.IsKeyboardFocusWithin);

            window.Close();
        });
    }

    [Fact]
    public void Tab_at_the_end_of_a_row_wraps_to_the_next_rows_first_cell_and_Shift_Tab_back()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = ShowWithControlBefore(editor);
            int lastColumn = BoardWorkbookSchema.Components.ColumnOrder.Count - 1;

            editor.SelectCell(Row(editor, "U1"), lastColumn);
            editor.GetControl<DataGrid>("TableGrid").Focus();

            Press(window, Avalonia.Input.Key.Tab);
            Assert.Same(Row(editor, "U2"), editor.CurrentRow);
            Assert.Equal(0, editor.CurrentCell!.ColumnIndex);

            Press(window, Avalonia.Input.Key.Tab, Avalonia.Input.RawInputModifiers.Shift);
            Assert.Same(Row(editor, "U1"), editor.CurrentRow);
            Assert.Equal(lastColumn, editor.CurrentCell!.ColumnIndex);

            window.Close();
        });
    }

    [Fact]
    public void Tab_while_typing_in_a_cell_keeps_what_was_typed_and_moves_on()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Window window = ShowWithControlBefore(editor);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            grid.Focus();
            window.KeyTextInput("CPU");
            Dispatcher.UIThread.RunJobs();

            Press(window, Avalonia.Input.Key.Tab);

            Assert.Equal("CPU", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);
            Assert.Equal(Column(BoardWorkbookSchema.ColFriendlyName) + 1, editor.CurrentCell!.ColumnIndex);
            Assert.True(grid.IsKeyboardFocusWithin);

            window.Close();
        });
    }

    [Fact]
    public void The_table_uses_a_smaller_font_and_fixed_row_height_to_fit_more_rows()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            Assert.Equal(12, grid.FontSize);
            Assert.Equal(26, grid.RowHeight);

            // The grid's own FontSize reaches only its headers: a text column carries its own size
            // for its cells, so every column must say 12 too - the first version set the grid's
            // alone and the cells did not shrink at all.
            Assert.All(grid.Columns.OfType<DataGridTextColumn>(), column => Assert.Equal(12, column.FontSize));
        });
    }

    [Fact]
    public void The_colour_key_counts_the_added_modified_and_deleted_rows_of_the_sheet_on_screen()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "CPU 6510"), Component("U9"), Component("U10")),
                published: Board(Component("U1", "CPU"), Component("U2")));

            Assert.Equal("2", editor.GetControl<TextBlock>("AddedCountText").Text);
            Assert.Equal("1", editor.GetControl<TextBlock>("ModifiedCountText").Text);
            Assert.Equal("1", editor.GetControl<TextBlock>("DeletedCountText").Text);

            // And they follow an edit.
            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "CPU";
            editor.RefreshPendingForTests();

            Assert.Equal("0", editor.GetControl<TextBlock>("ModifiedCountText").Text);
        });
    }

    [Fact]
    public void Only_changes_hides_unchanged_rows_and_switches_moving_off()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "CPU 6510"), Component("U3"), Component("U4")),
                published: Board(Component("U1", "CPU"), Component("U2"), Component("U3")));
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            editor.OnlyChanges = true;

            List<string> shown = ((System.Collections.IEnumerable)grid.ItemsSource!).Cast<BoardTableRow>()
                .Select(row => row.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text)
                .ToList();

            // U1 modified, U2 deleted, U4 added - and U3, unchanged, gone.
            Assert.Equal(["U1", "U2", "U4"], shown);
            Assert.DoesNotContain("RowsDraggable", grid.Classes);
            Assert.True(editor.GetControl<CheckBox>("OnlyChangesCheckBox").IsChecked);

            // Alt+Up does nothing meanwhile, either.
            editor.SelectCell(Row(editor, "U4"), 0);
            editor.MoveRowUp();
            Assert.Equal(["U1", "U3", "U4"], Components(editor).Rows.Where(row => !row.IsDeleted).Select(row => row.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text));

            editor.OnlyChanges = false;

            Assert.Equal(4, ((System.Collections.IEnumerable)grid.ItemsSource!).Cast<BoardTableRow>().Count());
            Assert.Contains("RowsDraggable", grid.Classes);
        });
    }

    [Fact]
    public void Under_only_changes_a_row_whose_change_is_reverted_drops_out_of_view()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU 6510")), published: Board(Component("U1", "CPU")));
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            editor.OnlyChanges = true;
            Assert.Single(((System.Collections.IEnumerable)grid.ItemsSource!).Cast<BoardTableRow>());

            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "CPU";
            editor.RefreshPendingForTests();

            Assert.Empty(((System.Collections.IEnumerable)grid.ItemsSource!).Cast<BoardTableRow>());
        });
    }

    [Fact]
    public void After_saving_a_placed_regional_variant_the_cursor_is_on_it_where_it_went()
    {
        // The project owner's case: the row seemed to vanish on save. It now goes beside its twin,
        // and the cursor follows it there rather than staying on whatever row took its old place.
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                Components =
                [
                    new ComponentEntry { BoardLabel = "U1", TechnicalNameOrValue = "x", Category = "IC", Region = "PAL" },
                    new ComponentEntry { BoardLabel = "U2", TechnicalNameOrValue = "x", Category = "IC" },
                    new ComponentEntry { BoardLabel = "C1", TechnicalNameOrValue = "x", Category = "Capacitor" },
                ],
            };
            BoardTableEditor editor = this.OpenEditor(board, published: board);
            Window window = Show(editor);

            BoardTableRow inserted = Components(editor).InsertRow(Row(editor, "C1"));
            inserted.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text = "U1";
            inserted.Cells[Column(BoardWorkbookSchema.ColRegion)].Text = "NTSC";
            editor.RefreshPendingForTests();
            editor.SelectCell(inserted, Column(BoardWorkbookSchema.ColRegion));

            Assert.Equal(BoardTableRowState.Added, inserted.State);
            Assert.Equal(DraftWorkbookEditOutcome.Saved, editor.Save());

            Assert.Equal(["U1", "U1", "U2", "C1"], Labels(editor));
            Assert.Equal("NTSC", editor.CurrentRow!.Cells[Column(BoardWorkbookSchema.ColRegion)].Text);
            Assert.Same(Components(editor).Rows[1], editor.CurrentRow);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Round 4 (2026-09-24)

    private static List<string> ShownLabels(BoardTableEditor editor) =>
        ((System.Collections.IEnumerable)editor.GetControl<DataGrid>("TableGrid").ItemsSource!).Cast<BoardTableRow>()
            .Select(row => row.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text)
            .ToList();

    [Fact]
    public void Show_changes_only_still_applies_after_switching_to_another_sheet_and_back()
    {
        // Reported: with the box ticked, leaving the sheet and coming back showed every row again
        // while the box still said it was ticked. Switched through the real TABS, in a shown window.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "CPU 6510"), Component("U3")),
                published: Board(Component("U1", "CPU"), Component("U3")));
            Window window = Show(editor);

            editor.OnlyChanges = true;
            Assert.Equal(["U1"], ShownLabels(editor));

            SheetTabs(editor).SelectedItem = SheetTabs(editor).Items.OfType<TabItem>()
                .First(tab => tab.Tag is BoardTableSheet { Name: BoardWorkbookSchema.SheetBoardSchematics });
            Dispatcher.UIThread.RunJobs();
            SheetTabs(editor).SelectedItem = SheetTabs(editor).Items.OfType<TabItem>()
                .First(tab => ReferenceEquals(tab.Tag, Components(editor)));
            Dispatcher.UIThread.RunJobs();

            Assert.Same(Components(editor), editor.CurrentSheet);
            Assert.True(editor.GetControl<CheckBox>("OnlyChangesCheckBox").IsChecked);
            Assert.Equal(["U1"], ShownLabels(editor));

            window.Close();
        });
    }

    // The pill a count sits in: its nearest Border carrying the LegendPill class.
    private static Border PillOf(BoardTableEditor editor, string countName) =>
        editor.GetControl<TextBlock>(countName).GetVisualAncestors().OfType<Border>()
            .First(border => border.Classes.Contains("LegendPill"));

    [Theory]
    [InlineData("AddedCountText", "Added", "BoardTable_Added_Bg")]
    [InlineData("ModifiedCountText", "Modified", "BoardTable_Modified_Bg")]
    [InlineData("DeletedCountText", "Deleted", "BoardTable_Deleted_Bg")]
    [InlineData("FlaggedCountText", "Flagged", "BoardTable_Flagged_Bg")]
    public void Each_count_shares_ONE_pill_with_its_own_word_in_its_rows_colour(string countName, string word, string colourKey)
    {
        // Reported: a separate badge before each word, with even gaps all along, read as easily
        // "Added 1" as "2 Added". Now the count and its word are one pill in the wash its rows are
        // painted in - and only the NUMBER is bold, as on the Workbooks tab's counted pills.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: Board(Component("U1")));
            Window window = Show(editor);

            Border pill = PillOf(editor, countName);
            List<TextBlock> texts = pill.GetVisualDescendants().OfType<TextBlock>().ToList();

            Assert.Equal([countName, null], texts.Select(text => text.Name));
            Assert.Equal(word, texts[1].Text);
            Assert.Equal(FontWeight.Bold, texts[0].FontWeight);
            Assert.NotEqual(FontWeight.Bold, texts[1].FontWeight);
            Assert.Equal(ThemeColor(colourKey), (pill.Background as ISolidColorBrush)?.Color);

            window.Close();
        });
    }

    [Fact]
    public void The_four_pills_are_four_separate_pills()
    {
        // Each pill holds its own count - no two counts inside the same one.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: Board(Component("U1")));

            List<Border> pills = new[] { "AddedCountText", "ModifiedCountText", "DeletedCountText", "FlaggedCountText" }
                .Select(name => PillOf(editor, name))
                .ToList();

            Assert.Equal(4, pills.Distinct().Count());
        });
    }

    [Fact]
    public void A_duplicate_row_is_painted_violet_counted_as_flagged_and_left_out_of_the_tab_number()
    {
        // Owner request: the "!" row should be visualised like the other three, not only
        // told by a character at its start.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1"), Component("U1", "Second U1"), Component("U9")),
                published: Board(Component("U1")));
            Window window = Show(editor);

            BoardTableRow duplicate = Components(editor).Rows.Single(row =>
                row.Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text == "Second U1");

            Assert.Equal("!", duplicate.Marker);
            Assert.Equal(
                ThemeColor("BoardTable_Flagged_Bg"),
                (CellOnScreen(window, duplicate, Column(BoardWorkbookSchema.ColFriendlyName)).Background as ISolidColorBrush)?.Color);

            Assert.Equal("1", editor.GetControl<TextBlock>("FlaggedCountText").Text);
            Assert.Equal("1", editor.GetControl<TextBlock>("AddedCountText").Text);

            // The tab says one change (U9), exactly as the draft row's own count does.
            Assert.Equal("Components (1)", SheetTab(editor, "Components (1)").Header);

            // And "Show changes only" keeps it - it is the contributor's to look at.
            editor.OnlyChanges = true;
            Assert.Contains(duplicate, ((System.Collections.IEnumerable)editor.GetControl<DataGrid>("TableGrid").ItemsSource!).Cast<BoardTableRow>());

            window.Close();
        });
    }

    [Fact]
    public void The_filter_box_reads_Show_changes_only()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: Board(Component("U1")));

            Assert.Equal("Show changes only", editor.GetControl<CheckBox>("OnlyChangesCheckBox").Content);
        });
    }

    // ------------------------------------------------------------------ Round 5 (2026-09-24)

    // A real click, so the grid really selects the cell - SelectCell alone moves the cursor but
    // leaves the cell unselected, and the selected look is what is being tested.
    private static void Click(Window window, Control target)
    {
        Point centre = target.TranslatePoint(new Point(target.Bounds.Width / 2, target.Bounds.Height / 2), window)!.Value;
        window.MouseDown(centre, Avalonia.Input.MouseButton.Left);
        window.MouseUp(centre, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    private static T TemplatePart<T>(Control owner, string name)
        where T : Control =>
        owner.GetVisualDescendants().OfType<T>().First(part => part.Name == name);

    [Fact]
    public void The_current_cell_is_not_filled_but_framed_by_a_2px_dashed_red_line()
    {
        // Reported: the grid filled the selected cell with its accent green, which could be taken
        // for "added". Asked for instead: no fill, and a 2px dashed red border.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2")), published: Board(Component("U1", "CPU"), Component("U2")));
            Window window = Show(editor);

            DataGridCell cell = CellOnScreen(window, Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            Click(window, cell);

            Assert.Contains(":selected", cell.Classes);
            Assert.Contains(":current", cell.Classes);
            Assert.Equal(Colors.Transparent, (cell.Background as ISolidColorBrush)?.Color);

            Avalonia.Controls.Shapes.Rectangle frame = TemplatePart<Avalonia.Controls.Shapes.Rectangle>(cell, "CurrencyVisual");
            Assert.True(frame.IsEffectivelyVisible);
            Assert.Equal(2, frame.StrokeThickness);
            Assert.Equal(ThemeColor("BoardTable_CurrentCell_Border"), (frame.Stroke as ISolidColorBrush)?.Color);
            Assert.NotNull(frame.StrokeDashArray);
            Assert.NotEmpty(frame.StrokeDashArray);

            // The grid's own solid frame is gone, not drawn underneath.
            Assert.False(TemplatePart<Grid>(cell, "FocusVisual").IsEffectivelyVisible);

            window.Close();
        });
    }

    [Fact]
    public void A_selected_changed_cell_keeps_its_orange_under_the_frame()
    {
        // The accent fill used to cover the cell's own colour while it was selected, so the one
        // cell being looked at was the one whose state could not be seen.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU 6510")), published: Board(Component("U1", "CPU")));
            Window window = Show(editor);

            DataGridCell cell = CellOnScreen(window, Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            Click(window, cell);

            Assert.Contains(":selected", cell.Classes);
            Assert.Equal(ThemeColor("BoardTable_Modified_Bg"), (cell.Background as ISolidColorBrush)?.Color);

            window.Close();
        });
    }

    [Fact]
    public void The_grids_selection_outline_and_fill_handle_are_hidden_and_the_handle_cannot_be_grabbed()
    {
        // The fill handle drags one value down a whole run of cells - a gesture this table was
        // never built or tested for - so it must not merely be invisible but unclickable too.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = Show(editor);
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            Click(window, CellOnScreen(window, Row(editor, "U1"), 0));

            Border outline = TemplatePart<Border>(grid, "PART_SelectionOutline");
            Border fillHandle = TemplatePart<Border>(grid, "PART_FillHandle");

            Assert.Equal(0, outline.Opacity);
            Assert.Equal(0, fillHandle.Opacity);
            Assert.False(fillHandle.IsHitTestVisible);

            window.Close();
        });
    }

    [Fact]
    public void The_row_drag_grip_is_centred_in_the_row_header()
    {
        // Reported as left-aligned: the template puts it in a 16px first column of the header's
        // two, so it sat centred in that column rather than in the header.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Window window = Show(editor);

            Avalonia.Controls.Primitives.DataGridRowHeader header = window.GetVisualDescendants().OfType<Avalonia.Controls.Primitives.DataGridRowHeader>()
                .First(candidate => ReferenceEquals(candidate.DataContext, Row(editor, "U1")));
            Border grip = TemplatePart<Border>(header, "DragGrip");

            double gripCentre = grip.TranslatePoint(new Point(grip.Bounds.Width / 2, 0), header)!.Value.X;

            Assert.InRange(gripCentre, (header.Bounds.Width / 2) - 1, (header.Bounds.Width / 2) + 1);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Undo / redo

    private static void PressCtrl(Window window, Avalonia.Input.Key key, bool shift = false)
    {
        Avalonia.Input.RawInputModifiers modifiers = Avalonia.Input.RawInputModifiers.Control;
        if (shift)
        {
            modifiers |= Avalonia.Input.RawInputModifiers.Shift;
        }

        window.KeyPress(key, modifiers, key == Avalonia.Input.Key.Z ? Avalonia.Input.PhysicalKey.Z : Avalonia.Input.PhysicalKey.Y, keySymbol: null);
        Dispatcher.UIThread.RunJobs();
    }

    // Types into the current cell and commits with Enter, as a contributor would.
    private static void TypeAndCommit(Window window, BoardTableEditor editor, string text)
    {
        editor.GetControl<DataGrid>("TableGrid").Focus();
        window.KeyTextInput(text);
        Dispatcher.UIThread.RunJobs();
        Press(window, Avalonia.Input.Key.Enter);
        editor.RefreshPendingForTests();
    }

    [Fact]
    public void Ctrl_Z_undoes_a_typed_cell_and_Ctrl_Y_and_Ctrl_Shift_Z_redo_it()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2")), published: Board(Component("U1", "CPU"), Component("U2")));
            Window window = ShowWithControlBefore(editor);
            BoardTableCell cell = Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)];

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            TypeAndCommit(window, editor, "6510");
            Assert.Equal("6510", cell.Text);

            PressCtrl(window, Avalonia.Input.Key.Z);
            Assert.Equal("CPU", cell.Text);
            Assert.Equal(BoardTableCellState.Unchanged, cell.State);

            // The cursor is back on the cell that changed, not wherever Enter had moved it.
            Assert.Same(cell, editor.CurrentCell);

            PressCtrl(window, Avalonia.Input.Key.Y);
            Assert.Equal("6510", cell.Text);

            PressCtrl(window, Avalonia.Input.Key.Z);
            PressCtrl(window, Avalonia.Input.Key.Z, shift: true);
            Assert.Equal("6510", cell.Text);

            window.Close();
        });
    }

    [Fact]
    public void Typing_a_word_into_a_cell_is_ONE_undo_step_not_one_per_letter()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Window window = ShowWithControlBefore(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            TypeAndCommit(window, editor, "Processor");

            Assert.Equal(1, editor.SessionForTests!.Document.History.UndoCount);

            window.Close();
        });
    }

    [Fact]
    public void Ctrl_Z_while_typing_in_a_cell_belongs_to_the_cell_and_leaves_the_table_alone()
    {
        // As in Excel: inside a cell being edited, Ctrl+Z is the text box's own - taking back
        // an earlier, committed change instead would be a surprise.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: null);
            Window window = ShowWithControlBefore(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            TypeAndCommit(window, editor, "Earlier");

            editor.SelectCell(Row(editor, "U2"), Column(BoardWorkbookSchema.ColFriendlyName));
            editor.GetControl<DataGrid>("TableGrid").Focus();
            window.KeyTextInput("Now");
            Dispatcher.UIThread.RunJobs();

            PressCtrl(window, Avalonia.Input.Key.Z);

            Assert.Equal("Earlier", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);
            Assert.Equal(1, editor.SessionForTests!.Document.History.UndoCount);

            window.Close();
        });
    }

    [Fact]
    public void Undo_goes_back_to_the_sheet_where_the_change_was()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: Board(Component("U1", "CPU")));
            Window window = ShowWithControlBefore(editor);
            BoardTableCell cell = Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)];

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            TypeAndCommit(window, editor, "Changed");

            BoardTableSheet credits = editor.SessionForTests!.Document.FindSheet(BoardWorkbookSchema.SheetCredits)!;
            editor.SelectSheet(credits);
            editor.GetControl<DataGrid>("TableGrid").Focus();

            PressCtrl(window, Avalonia.Input.Key.Z);

            Assert.Same(Components(editor), editor.CurrentSheet);
            Assert.Equal("CPU", cell.Text);
            Assert.Same(cell, editor.CurrentCell);

            window.Close();
        });
    }

    [Fact]
    public void Ctrl_Z_works_right_after_a_toolbar_button_was_used()
    {
        // "Delete row", then Ctrl+Z. The delete leaves the cursor on the red ghost, which cannot be
        // deleted, so the button disables itself - and a focused button that disables drops the
        // focus to nothing, where Ctrl+Z reached no handler at all. The button now hands the
        // keyboard back to the table.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1"), Component("U2")), published: Board(Component("U1"), Component("U2")));
            Window window = ShowWithControlBefore(editor);
            BoardTableRow u2 = Row(editor, "U2");

            editor.SelectCell(u2, 0);
            Button delete = editor.GetControl<Button>("DeleteRowButton");
            delete.Focus();
            delete.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.DoesNotContain(u2, Components(editor).Rows);
            Assert.False(delete.IsEnabled);
            Assert.True(editor.GetControl<DataGrid>("TableGrid").IsKeyboardFocusWithin);

            PressCtrl(window, Avalonia.Input.Key.Z);

            Assert.Contains(u2, Components(editor).Rows);
            Assert.Equal(0, Components(editor).DeletedCount);

            window.Close();
        });
    }

    [Fact]
    public void Ctrl_Z_works_with_the_focus_on_a_control_beside_the_table()
    {
        // Undo is handled on the whole editor, not only the grid: ticking "Show changes only"
        // leaves the focus on the check box, and Ctrl+Z from there must still undo.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: Board(Component("U1", "CPU")));
            Window window = ShowWithControlBefore(editor);
            BoardTableCell cell = Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)];

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            TypeAndCommit(window, editor, "Changed");

            CheckBox onlyChanges = editor.GetControl<CheckBox>("OnlyChangesCheckBox");
            Assert.True(onlyChanges.Focus());
            Dispatcher.UIThread.RunJobs();

            PressCtrl(window, Avalonia.Input.Key.Z);

            Assert.Equal("CPU", cell.Text);

            window.Close();
        });
    }

    [Fact]
    public void Undoing_every_change_greys_Save_changes_out_again()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: Board(Component("U1", "CPU")));
            Window window = ShowWithControlBefore(editor);
            Button save = editor.GetControl<Button>("SaveButton");

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            TypeAndCommit(window, editor, "Changed");
            Assert.True(save.IsEnabled);

            PressCtrl(window, Avalonia.Input.Key.Z);

            Assert.False(editor.HasUnsavedChanges);
            Assert.False(save.IsEnabled);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Round 6 (2026-09-24)

    [Fact]
    public void Show_changes_only_still_applies_after_switching_to_another_MAIN_tab_and_back()
    {
        // Reported: ticked, then Drafts -> Contribute -> Drafts showed every row again while the
        // box stayed ticked. Leaving a tab detaches its content and coming back re-attaches it,
        // which is what this does - through a real TabControl, as the main window's.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "CPU 6510"), Component("U3")),
                published: Board(Component("U1", "CPU"), Component("U3")));

            var tabs = new TabControl
            {
                ItemsSource = new[]
                {
                    new TabItem { Header = "Drafts", Content = editor },
                    new TabItem { Header = "Contribute", Content = new TextBlock { Text = "elsewhere" } },
                },
            };
            var window = new Window { Content = tabs, Width = 1300, Height = 600 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            editor.OnlyChanges = true;
            Assert.Equal(["U1"], ShownLabels(editor));

            tabs.SelectedIndex = 1;
            Dispatcher.UIThread.RunJobs();
            tabs.SelectedIndex = 0;
            Dispatcher.UIThread.RunJobs();

            Assert.True(editor.GetControl<CheckBox>("OnlyChangesCheckBox").IsChecked);
            Assert.Equal(["U1"], ShownLabels(editor));

            window.Close();
        });
    }

    [Fact]
    public void A_draft_with_nothing_published_never_inherits_a_hidden_show_changes_only()
    {
        // Reported as "the Board schematics sheet is empty although it has 3 schematic images":
        // the box was ticked on a published board's table, and the SAME editor then opened a
        // system with nothing published - where the box is hidden, every row is unchanged, and so
        // a filter still on hid every row with no way to see why.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1", "CPU 6510")),
                published: Board(Component("U1", "CPU")));
            editor.OnlyChanges = true;

            const string otherKey = "Manu5/Hw5/Board5/Data Hw5 Board5.xlsx";
            string otherPath = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, otherKey);
            Directory.CreateDirectory(Path.GetDirectoryName(otherPath)!);
            BoardWorkbookWriter.Write(otherPath, new BoardData
            {
                Schematics =
                [
                    new BoardSchematicEntry { SchematicName = "M1_1_PAL", SchematicImageFile = "Manu5/Hw5/Board5/M1_1_PAL.png" },
                    new BoardSchematicEntry { SchematicName = "M1_2_PAL", SchematicImageFile = "Manu5/Hw5/Board5/M1_2_PAL.png" },
                    new BoardSchematicEntry { SchematicName = "Gode sager", SchematicImageFile = "Manu5/Hw5/Board5/M1_3_PAL.png" },
                ],
            });

            Assert.True(editor.Load(this.DraftsRoot, otherKey, published: null));

            Assert.False(editor.GetControl<CheckBox>("OnlyChangesCheckBox").IsVisible);
            Assert.False(editor.OnlyChanges);
            Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, editor.CurrentSheet!.Name);
            Assert.Equal(3, ((System.Collections.IEnumerable)editor.GetControl<DataGrid>("TableGrid").ItemsSource!).Cast<BoardTableRow>().Count());
        });
    }

    // ------------------------------------------------------------------ Round 7 (2026-09-24)

    [Fact]
    public void A_pill_counting_nothing_fades_back_like_a_disabled_control()
    {
        // Owner request: "0 Added" should steal less attention than the kinds that are there.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "changed")), published: Board(Component("U1")));
            Window window = Show(editor);

            Assert.Equal(0.4, editor.GetControl<Border>("AddedPill").Opacity, 3);
            Assert.Equal(0.4, editor.GetControl<Border>("DeletedPill").Opacity, 3);
            Assert.Equal(0.4, editor.GetControl<Border>("FlaggedPill").Opacity, 3);
            Assert.Equal(1, editor.GetControl<Border>("ModifiedPill").Opacity, 3);

            // And it fades as soon as there is nothing left to count.
            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = string.Empty;
            editor.RefreshPendingForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(0.4, editor.GetControl<Border>("ModifiedPill").Opacity, 3);

            window.Close();
        });
    }

    [Fact]
    public void The_sheet_tabs_follow_the_workbook_with_Credits_last()
    {
        // Reported: "Important signals" came last, but every published workbook ends in "Credits".
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            List<string> headers = SheetTabs(editor).Items.OfType<TabItem>().Select(tab => (string)tab.Header!).ToList();

            Assert.Equal(BoardWorkbookSchema.SheetCredits, headers[^1]);
            Assert.Equal(BoardWorkbookSchema.SheetKiCadImportantSignals, headers[^2]);
        });
    }

    // ------------------------------------------------------------------ Round 8 (2026-09-24)

    [Fact]
    public void Moves_arriving_before_the_grid_has_redrawn_do_not_flicker_the_row_back()
    {
        // Reported: while dragging, the row flickered between two places. Pointer moves can arrive
        // before the grid has laid its rows out again after the previous move, so anything that
        // asks "which row is under the pointer" of the LIVE layout gets a stale answer and moves
        // the row back. Here several moves arrive with no layout pass in between.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(
                Board(Component("U1"), Component("U2"), Component("U3"), Component("U4"), Component("U5")),
                published: null);
            Window window = Show(editor);

            Point grip = CentreOf(window, HeaderOf(window, Row(editor, "U1")));
            Point overU3 = CentreOf(window, CellOnScreen(window, Row(editor, "U3"), Column(BoardWorkbookSchema.ColFriendlyName)));

            PressAt(window, grip);

            for (int i = 0; i < 6; i++)
            {
                window.MouseMove(overU3 + new Point(0, i % 3), Avalonia.Input.RawInputModifiers.LeftMouseButton);
            }

            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["U2", "U3", "U1", "U4", "U5"], Labels(editor));

            // And again after a redraw, still resting.
            DragTo(window, overU3 + new Point(0, 1));
            Assert.Equal(["U2", "U3", "U1", "U4", "U5"], Labels(editor));

            ReleaseAt(window, overU3);
            window.Close();
        });
    }

    [Fact]
    public void Holding_a_dragged_row_still_over_the_half_visible_bottom_row_does_not_run_away()
    {
        // THE REPORTED FLICKER, reproduced: on a sheet long enough to scroll, the first version
        // scrolled the moved row into view after every move. Over the half-visible bottom row that
        // scroll slid a NEW row under a pointer that had not moved, which was moved onto, which
        // scrolled again - the row ran down the sheet (or flickered) for as long as the mouse
        // twitched there. The target is now read off the layout frozen when the drag started.
        UiTest.Run(() =>
        {
            ComponentEntry[] many = Enumerable.Range(1, 40).Select(i => Component($"C{i}")).ToArray();
            BoardTableEditor editor = this.OpenEditor(Board(many), published: Board(many));
            var window = new Window { Content = editor, Width = 1200, Height = 432 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            // The lowest row on screen, only partly visible at this window height - and a point
            // INSIDE the rows area over it (below the rows area is the edge, where each move is
            // meant to carry the row one further).
            var presenter = window.GetVisualDescendants().OfType<Avalonia.Controls.Primitives.DataGridRowsPresenter>().Single();
            double rowsBottom = presenter.TranslatePoint(new Point(0, presenter.Bounds.Height), window)!.Value.Y;

            DataGridRow bottom = window.GetVisualDescendants().OfType<DataGridRow>()
                .Where(row => row.IsVisible && row.TranslatePoint(default, window)!.Value.Y < rowsBottom)
                .OrderBy(row => row.TranslatePoint(default, window)!.Value.Y)
                .Last();
            double bottomTop = bottom.TranslatePoint(default, window)!.Value.Y;
            string bottomLabel = ((BoardTableRow)bottom.DataContext!).Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text;

            Assert.True(bottomTop + bottom.Bounds.Height > rowsBottom, "the bottom row should be only partly visible");
            Assert.True(bottomTop + 5 < rowsBottom, "the point should be inside the rows area");

            Point grip = CentreOf(window, HeaderOf(window, Row(editor, "C3")));
            PressAt(window, grip);

            DragTo(window, new Point(grip.X + 200, bottomTop + 4));
            List<string> afterFirstMove = Labels(editor);
            Assert.Equal("C3", afterFirstMove[afterFirstMove.IndexOf(bottomLabel) + 1]);

            // The mouse twitches in place: nothing more may move.
            foreach (double dy in new[] { 4.0, 5.0, 3.0, 4.0, 2.0 })
            {
                DragTo(window, new Point(grip.X + 200, bottomTop + dy));
                Assert.Equal(afterFirstMove, Labels(editor));
            }

            ReleaseAt(window, grip);
            Assert.Equal(afterFirstMove, Labels(editor));
            window.Close();
        });
    }

    // ------------------------------------------------------------------ Watching the draft file (2026-09-24)

    [Fact]
    public void A_refused_save_raises_the_warning_bar_too()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: null);
            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "Unsaved";
            editor.RefreshPendingForTests();

            this.WriteDraft(Board(Component("U1", "From Excel")));

            Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, editor.Save());
            Assert.True(editor.GetControl<Border>("ChangedOnDiskBar").IsVisible);
            Assert.False(editor.GetControl<Button>("SaveButton").IsEnabled);
        });
    }

    [Fact]
    public void A_cell_being_typed_in_is_never_reloaded_away_by_an_outside_change()
    {
        // "Nothing unsaved" is not the whole truth while a cell's editor is open - the typing has
        // not reached the model. Reloading then would throw it away under the contributor's hands.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: null);
            Window window = ShowWithControlBefore(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColFriendlyName));
            editor.GetControl<DataGrid>("TableGrid").Focus();
            window.KeyTextInput("Typing");
            Dispatcher.UIThread.RunJobs();
            Assert.False(editor.HasUnsavedChanges);

            this.WriteDraft(Board(Component("U1", "From Excel")));
            editor.CheckDraftFile();

            // Not reloaded: the typing goes in when committed ...
            Press(window, Avalonia.Input.Key.Enter);
            editor.RefreshPendingForTests();
            Assert.Equal("Typing", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);

            // ... and the next check finds unsaved edits on a changed draft: the warning, not a reload.
            editor.CheckDraftFile();
            Assert.True(editor.GetControl<Border>("ChangedOnDiskBar").IsVisible);
            Assert.Equal("Typing", Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text);

            window.Close();
        });
    }

    [Fact]
    public void The_open_in_Excel_notice_shows_the_whole_time_the_lock_file_is_there_edits_or_not()
    {
        // Owner's choice: "edit it in one place at a time" matters most BEFORE any editing
        // starts, so the notice does not wait for the table to have unsaved edits.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Border notice = editor.GetControl<Border>("OpenElsewhereBar");
            string lockFile = Path.Combine(Path.GetDirectoryName(this.WorkbookPath)!, "~$" + Path.GetFileName(this.WorkbookPath));

            editor.CheckDraftFile();
            Assert.False(notice.IsVisible);

            File.WriteAllText(lockFile, "owner");
            editor.CheckDraftFile();
            Assert.True(notice.IsVisible);
            Assert.Contains("Edit it in one place at a time", editor.GetControl<TextBlock>("OpenElsewhereText").Text);

            // Still there with edits - and saving stays possible, since a lock file left behind by
            // a crashed Excel must not block the table for good.
            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "Edited";
            editor.RefreshPendingForTests();
            editor.CheckDraftFile();
            Assert.True(notice.IsVisible);
            Assert.True(editor.GetControl<Button>("SaveButton").IsEnabled);

            File.Delete(lockFile);
            editor.CheckDraftFile();
            Assert.False(notice.IsVisible);
        });
    }

    [Fact]
    public void A_table_opened_while_the_draft_is_already_open_in_Excel_says_so_at_once()
    {
        // Reported: launch, open the table with Excel already holding the draft - no notice. It is
        // now set when the table reads the file, not left to the first check two seconds later.
        UiTest.Run(() =>
        {
            this.WriteDraft(Board(Component("U1")));
            File.WriteAllText(
                Path.Combine(Path.GetDirectoryName(this.WorkbookPath)!, "~$" + Path.GetFileName(this.WorkbookPath)),
                "owner");

            var editor = new BoardTableEditor();
            Assert.True(editor.Load(this.DraftsRoot, BoardTableEditorTests.SystemKey, published: null));

            Assert.True(editor.GetControl<Border>("OpenElsewhereBar").IsVisible);
        });
    }

    [Fact]
    public void A_reload_caused_by_an_outside_change_is_announced_so_the_draft_row_and_board_follow()
    {
        // Reported (seen in a screenshot): after Excel changed the draft and the table reloaded,
        // the draft row still said "9 rows changed" while the tabs added up to 10, and the
        // component list still showed the old value. A save announces itself; this reload must too.
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")), published: null);
            int announced = 0;
            editor.ReloadedFromOutside += (_, _) => announced++;

            editor.CheckDraftFile();
            Assert.Equal(0, announced);

            this.WriteDraft(Board(Component("U1", "From Excel")));
            editor.CheckDraftFile();
            Assert.Equal(1, announced);

            // With unsaved edits there is no reload, so nothing to announce.
            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "Unsaved";
            editor.RefreshPendingForTests();
            this.WriteDraft(Board(Component("U1", "Again")));
            editor.CheckDraftFile();
            Assert.Equal(1, announced);
        });
    }

    [Fact]
    public void The_notices_are_roomy()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Window window = Show(editor);

            Assert.Equal(new Thickness(15, 11), editor.GetControl<Border>("ChangedOnDiskBar").Padding);
            Assert.Equal(new Thickness(15, 11), editor.GetControl<Border>("OpenElsewhereBar").Padding);

            window.Close();
        });
    }

    [Fact]
    public void The_draft_file_is_watched_only_while_the_table_is_on_screen()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Assert.False(editor.IsWatchingDraftFileForTests);

            Window window = Show(editor);
            Assert.True(editor.IsWatchingDraftFileForTests);

            // A tab switch detaches the editor, as closing the window does here.
            window.Close();
            Dispatcher.UIThread.RunJobs();
            Assert.False(editor.IsWatchingDraftFileForTests);
        });
    }

    [Fact]
    public void No_table_no_watching()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Window window = Show(editor);

            editor.Clear();

            Assert.False(editor.IsWatchingDraftFileForTests);
            window.Close();
        });
    }

    [Fact]
    public void Save_is_greyed_out_while_Excel_really_holds_the_draft_open_and_back_once_it_lets_go()
    {
        // Owner request: "if I cannot save before Excel is closed, the save button should be
        // disabled until the file gets closed". Held = lock file AND the workbook really held open,
        // so a lock file left behind by a crashed Excel does not block the table for good.
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File sharing is advisory outside Windows.");

        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1")), published: null);
            Button save = editor.GetControl<Button>("SaveButton");
            string lockFile = Path.Combine(Path.GetDirectoryName(this.WorkbookPath)!, "~$" + Path.GetFileName(this.WorkbookPath));

            Row(editor, "U1").Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "Edited";
            editor.RefreshPendingForTests();
            Assert.True(save.IsEnabled);

            File.WriteAllText(lockFile, "owner");

            using (new FileStream(this.WorkbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                editor.CheckDraftFile();
                Assert.True(editor.GetControl<Border>("OpenElsewhereBar").IsVisible);
                Assert.False(save.IsEnabled);
            }

            // Excel closed the file (its lock file not yet gone, as after a crash): Save is back.
            editor.CheckDraftFile();
            Assert.True(save.IsEnabled);

            File.Delete(lockFile);
            editor.CheckDraftFile();
            Assert.True(save.IsEnabled);
            Assert.False(editor.GetControl<Border>("OpenElsewhereBar").IsVisible);
        });
    }
    // ###########################################################################################
    // DOCUMENT MODE - the maintainer application's table (2026-09-25). The same editor, opened on a
    // document with no draft file behind it: Save hands the document to the host, nothing is
    // written anywhere, and the controls that only make sense for a draft file stay out of sight.
    // ###########################################################################################
    private static BoardTableEditor OpenDocument(BoardData submitted, BoardData? published)
    {
        var editor = new BoardTableEditor();
        editor.Open(BoardTableDocument.Create(published, submitted));
        return editor;
    }

    [Fact]
    public void A_document_opens_with_no_file_behind_it_and_no_Reload()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = OpenDocument(Board(Component("U1")), published: Board(Component("U1")));
            Window window = Show(editor);

            Assert.True(editor.HasTable);
            Assert.False(editor.IsFileBacked);
            Assert.False(editor.GetControl<Button>("ReloadButton").IsVisible);
            Assert.False(editor.GetControl<Border>("ChangedOnDiskBar").IsVisible);
            Assert.False(editor.GetControl<Border>("OpenElsewhereBar").IsVisible);

            window.Close();
        });
    }

    // Save in document mode hands the edited document to the host - the maintainer application
    // sends it to the server - and writes nothing itself.
    [Fact]
    public void Save_in_a_document_asks_the_host_and_hands_it_the_edited_document()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = OpenDocument(Board(Component("U1")), published: Board(Component("U1")));
            Window window = Show(editor);
            BoardTableSheet components = editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
            editor.SelectSheet(components);

            int requests = 0;
            editor.SaveRequested += (_, _) => requests++;

            components.Rows.Single().Cells[Column(BoardWorkbookSchema.ColFriendlyName)].Text = "CPU 6510";
            editor.RefreshPendingForTests();

            Button save = editor.GetControl<Button>("SaveButton");
            Assert.True(save.IsEnabled);

            save.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            Assert.Equal(1, requests);

            BoardData edited = editor.CommitAndGetDocument()!.ApplyTo(Board(Component("U1")));
            Assert.Equal("CPU 6510", Assert.Single(edited.Components).FriendlyName);

            window.Close();
        });
    }

    // After the host saved, it opens the saved state again - on the same sheet, with nothing
    // unsaved - and says so.
    [Fact]
    public void Opening_the_saved_state_again_keeps_the_sheet_and_clears_unsaved()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = OpenDocument(Board(Component("U1")), published: Board(Component("U1")));
            Window window = Show(editor);
            editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!);

            editor.Open(BoardTableDocument.Create(Board(Component("U1")), Board(Component("U1", "CPU 6510"))), "Saved.");

            Assert.Equal(BoardWorkbookSchema.SheetComponents, editor.CurrentSheet!.Name);
            Assert.False(editor.HasUnsavedChanges);
            Assert.Equal("Saved.", editor.GetControl<TextBlock>("StatusText").Text);

            window.Close();
        });
    }
    // A maintainer opening a submission lands on what changed, not on the first sheet every time.
    [Fact]
    public void A_document_opens_on_the_first_sheet_with_a_change()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = OpenDocument(Board(Component("U1", "CPU 6510")), published: Board(Component("U1")));

            Assert.Equal(BoardWorkbookSchema.SheetComponents, editor.CurrentSheet!.Name);
        });
    }
}
