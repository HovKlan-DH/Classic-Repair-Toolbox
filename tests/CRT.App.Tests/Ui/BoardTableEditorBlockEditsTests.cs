using System.IO;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The table's block edits on screen - copying and pasting blocks, fill down, and the replace bar
// (CRT 4.0, 2026-10-25: the plan's step 0, "the table view gains block copy and paste, fill down,
// and find & replace"; BoardTableEditor.BlockEdits.cs).
//
// The rules are CRT.Data's (BoardTableSheetTests, BoardTableDocumentTests, BoardTableClipboardTests);
// these hold the editor to handing them the right cells - the selection, the rows shown, whether
// rows may be added - and the keys and buttons to reaching them. Keys go through the window, as a
// contributor's do.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableEditorBlockEditsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private const string BoardKey = "Commodore/C64/250407/board.json";

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, TechnicalNameOrValue = $"part {label}" };

    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    private BoardTableEditor OpenEditor(BoardData draft)
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardTableEditorBlockEditsTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, draft);

        var editor = new BoardTableEditor();
        Assert.True(editor.Load(this.DraftsRoot, BoardTableEditorBlockEditsTests.BoardKey, published: null));
        editor.SelectSheet(Components(editor));

        return editor;
    }

    private static BoardTableSheet Components(BoardTableEditor editor) =>
        editor.DocumentForTests!.FindSheet(BoardWorkbookSchema.SheetComponents)!;

    private static int Column(string name) =>
        BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(name);

    private static int NameColumn => Column(BoardWorkbookSchema.ColFriendlyName);

    private static string Label(BoardTableRow row) => row.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text;

    private static string Name(BoardTableRow row) => row.Cells[NameColumn].Text;

    private static BoardTableRow Row(BoardTableEditor editor, string label) =>
        Components(editor).Rows.Single(row => !row.IsDeleted && Label(row) == label);

    private static Window Show(BoardTableEditor editor)
    {
        var window = new Window { Content = editor, Width = 1200, Height = 600 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    private static void Press(Window window, Key key, RawInputModifiers modifiers, PhysicalKey physical)
    {
        window.KeyPress(key, modifiers, physical, keySymbol: null);
        Dispatcher.UIThread.RunJobs();
    }

    // The cells of `rows` in one column selected, the cursor on the first - as a Shift+click down
    // the column leaves them.
    private static void SelectColumnOf(BoardTableEditor editor, int column, params BoardTableRow[] rows)
    {
        editor.SelectCell(rows[0], column);

        DataGrid grid = editor.GetControl<DataGrid>("TableGrid");
        DataGridColumn gridColumn = grid.Columns.Single(c => c.Tag is int index && index == column);
        var view = (Avalonia.Collections.DataGridCollectionView)grid.ItemsSource!;

        grid.SelectedCells = rows
            .Select(row => new DataGridCellInfo(row, gridColumn, view.IndexOf(row), gridColumn.DisplayIndex, true))
            .ToList();

        Dispatcher.UIThread.RunJobs();
    }

    private static string Status(BoardTableEditor editor) => editor.GetControl<TextBlock>("StatusText").Text ?? string.Empty;

    // ------------------------------------------------------------------ Copy and paste

    // What went in is selected afterwards - so it is in view, and copying straight away gives the
    // very block back.
    [Fact]
    public void A_pasted_block_is_selected_and_copies_back_as_the_same_block()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")));
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U2"), NameColumn);

            const string block = "VIC II\tMOS 6569\r\nSID 6581\tMOS 6581";
            Assert.True(editor.Paste(block + "\r\n"));

            Assert.Equal(["U1:CPU", "U2:VIC II", "U3:SID 6581"], Components(editor).Rows.Select(row => $"{Label(row)}:{Name(row)}"));
            Assert.Equal(4, editor.SelectedTableCells().Count);
            Assert.Equal(block, editor.CopySelectionText());

            window.Close();
        });
    }

    [Fact]
    public void A_block_longer_than_the_sheet_adds_rows_at_the_end()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")));
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), Column(BoardWorkbookSchema.ColBoardLabel));

            Assert.True(editor.Paste("U1\tCPU\r\nU8\tPLA\r\nU9\tROM\r\n"));

            Assert.Equal(["U1", "U8", "U9"], Components(editor).Rows.Select(Label));
            Assert.Contains("2 of them were added at the end of the sheet.", Status(editor));
            Assert.True(editor.HasUnsavedChanges);

            window.Close();
        });
    }

    // Narrowed by a search, a row added at the end would be out of sight: none is, and the status
    // says what was left out.
    [Fact]
    public void While_a_search_narrows_the_table_a_paste_adds_no_rows_and_says_so()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC")));
            Window window = Show(editor);

            editor.SearchText = "U1";
            editor.SelectCell(Row(editor, "U1"), NameColumn);

            Assert.True(editor.Paste("one\r\ntwo\r\nthree\r\n"));

            Assert.Equal("one", Name(Row(editor, "U1")));
            Assert.Equal("VIC", Name(Row(editor, "U2")));
            Assert.Equal(2, Components(editor).Rows.Count);
            Assert.Contains("2 rows were left out", Status(editor));

            window.Close();
        });
    }

    // ###########################################################################################
    // A paste or a fill that can change nothing says why (code review, 2026-10-25): a clipboard
    // holding a picture, or nothing, did nothing at all; red rows read as "already held it".
    // ###########################################################################################
    private BoardTableEditor OpenWithRedRows()
    {
        BoardData published = Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID"), Component("U4", "PLA"));
        BoardData draft = Board(Component("U1", "CPU"), Component("U4", "PLA"));

        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardTableEditorBlockEditsTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, draft);

        var editor = new BoardTableEditor();
        Assert.True(editor.Load(this.DraftsRoot, BoardTableEditorBlockEditsTests.BoardKey, published));
        editor.SelectSheet(Components(editor));

        return editor;
    }

    private static BoardTableRow RedRow(BoardTableEditor editor, string label) =>
        Components(editor).Rows.Single(row => row.IsDeleted && Label(row) == label);

    [Fact]
    public void A_clipboard_with_no_text_says_so()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")));
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), NameColumn);

            Assert.False(editor.Paste(null));
            Assert.Equal(BoardTableEditWording.NothingToPaste, Status(editor));
            Assert.Equal("CPU", Name(Row(editor, "U1")));

            window.Close();
        });
    }

    [Fact]
    public void One_value_pasted_onto_red_rows_alone_says_they_are_deleted()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenWithRedRows();
            Window window = Show(editor);

            SelectColumnOf(editor, NameColumn, RedRow(editor, "U2"), RedRow(editor, "U3"));

            Assert.False(editor.Paste("chip"));
            Assert.Equal(BoardTableEditWording.PasteOnDeletedRows, Status(editor));
            Assert.False(editor.HasUnsavedChanges);

            window.Close();
        });
    }

    [Fact]
    public void Ctrl_D_on_a_red_row_says_it_is_deleted()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenWithRedRows();
            Window window = Show(editor);

            editor.SelectCell(RedRow(editor, "U2"), NameColumn);
            editor.FocusGrid();

            Press(window, Key.D, RawInputModifiers.Control, PhysicalKey.D);

            Assert.Equal(BoardTableEditWording.FillOnDeletedRows, Status(editor));
            Assert.False(editor.HasUnsavedChanges);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Fill down

    [Fact]
    public void Ctrl_D_on_one_row_fills_it_from_the_row_above()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC")));
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U2"), NameColumn);
            editor.FocusGrid();

            Press(window, Key.D, RawInputModifiers.Control, PhysicalKey.D);

            Assert.Equal("CPU", Name(Row(editor, "U2")));
            Assert.Equal("Filled 1 cell with the value at the top. Ctrl+Z undoes it.", Status(editor));

            window.Close();
        });
    }

    [Fact]
    public void Fill_down_fills_the_selected_rows_from_the_top_one_and_one_undo_takes_it_back()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")));
            Window window = Show(editor);

            SelectColumnOf(editor, NameColumn, Row(editor, "U1"), Row(editor, "U2"), Row(editor, "U3"));

            Button fillDown = editor.GetControl<Button>("FillDownButton");
            Assert.True(fillDown.IsEnabled);

            fillDown.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["CPU", "CPU", "CPU"], Components(editor).Rows.Select(Name));

            editor.Undo();
            Assert.Equal(["CPU", "VIC", "SID"], Components(editor).Rows.Select(Name));

            window.Close();
        });
    }

    // On the first row with nothing above it, there is nothing to fill from.
    [Fact]
    public void Fill_down_is_off_on_the_top_row_alone()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC")));
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), NameColumn);
            Assert.False(editor.GetControl<Button>("FillDownButton").IsEnabled);

            editor.SelectCell(Row(editor, "U2"), NameColumn);
            Assert.True(editor.GetControl<Button>("FillDownButton").IsEnabled);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Find & replace

    [Fact]
    public void Replace_all_from_the_bar_replaces_in_the_rows_shown_and_one_undo_takes_it_back()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC")));
            Window window = Show(editor);

            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "PART";
            editor.GetControl<TextBox>("ReplaceWithBox").Text = "chip";
            editor.GetControl<Button>("ReplaceAllButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            int part = Column(BoardWorkbookSchema.ColTechnicalNameOrValue);
            Assert.Equal(["chip U1", "chip U2"], Components(editor).Rows.Select(row => row.Cells[part].Text));
            Assert.Equal("Replaced [PART] 2 times in 2 cells. Ctrl+Z undoes it.", Status(editor));

            editor.Undo();
            Assert.Equal(["part U1", "part U2"], Components(editor).Rows.Select(row => row.Cells[part].Text));

            window.Close();
        });
    }

    [Fact]
    public void Ctrl_H_opens_the_replace_bar_on_its_Find_box_and_Escape_closes_it()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")));
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), NameColumn);
            editor.FocusGrid();

            Press(window, Key.H, RawInputModifiers.Control, PhysicalKey.H);

            Assert.True(editor.IsReplaceBarOpen);
            Assert.True(editor.GetControl<TextBox>("ReplaceFindBox").IsFocused);

            Press(window, Key.Escape, RawInputModifiers.None, PhysicalKey.Escape);

            Assert.False(editor.IsReplaceBarOpen);

            window.Close();
        });
    }

    [Fact]
    public void Ctrl_F_puts_the_cursor_in_the_search_box()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU")));
            Window window = Show(editor);

            editor.SelectCell(Row(editor, "U1"), NameColumn);
            editor.FocusGrid();

            Press(window, Key.F, RawInputModifiers.Control, PhysicalKey.F);

            Assert.True(editor.GetControl<TextBox>("SearchBox").IsFocused);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Read-only

    // The Boards screen's table of a board the account may not change: copying stays, every
    // change is gone.
    [Fact]
    public void A_read_only_table_offers_no_fill_down_or_replace_and_pastes_nothing()
    {
        UiTest.Run(() =>
        {
            BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("U1", "CPU"), Component("U2", "VIC")));
            BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

            var editor = new BoardTableEditor();
            editor.Open(document);
            editor.IsReadOnly = true;
            Window window = Show(editor);

            editor.SelectSheet(components);
            BoardTableRow u2 = components.Rows.Single(row => Label(row) == "U2");
            editor.SelectCell(u2, NameColumn);
            editor.FocusGrid();

            Assert.False(editor.GetControl<Button>("FillDownButton").IsVisible);
            Assert.False(editor.GetControl<Button>("ReplaceButton").IsVisible);

            Assert.False(editor.Paste("pasted"));
            Press(window, Key.D, RawInputModifiers.Control, PhysicalKey.D);
            Press(window, Key.H, RawInputModifiers.Control, PhysicalKey.H);

            Assert.Equal("VIC", Name(u2));
            Assert.False(editor.IsReplaceBarOpen);
            Assert.Equal("VIC", editor.CopySelectionText());

            window.Close();
        });
    }
}
