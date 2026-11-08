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

    // ------------------------------------------------------------------ No fill down

    // ###########################################################################################
    // "Fill down" went (owner request, 2026-11-07: "I do not understand what "Fill down" does ... please remove this
    // functionality entirely"): no button, and Ctrl+D changes nothing.
    // ###########################################################################################
    [Fact]
    public void There_is_no_fill_down_and_Ctrl_D_changes_nothing()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC")));
            Window window = Show(editor);

            Assert.Null(editor.FindControl<Button>("FillDownButton"));

            editor.SelectCell(Row(editor, "U2"), NameColumn);
            editor.FocusGrid();

            Press(window, Key.D, RawInputModifiers.Control, PhysicalKey.D);

            Assert.Equal("VIC", Name(Row(editor, "U2")));
            Assert.False(editor.HasUnsavedChanges);

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

    // ###########################################################################################
    // "Replace" on a table filtered by the search box starts "Find" with that text, the cursor after it and nothing
    // selected (owner request, 2026-11-07: 'it should not "mark all text" in the "Find" input field. It should just place
    // (with focus) the cursor at the end of the text, which should match the filtering text').
    // ###########################################################################################
    [Fact]
    public void Replace_starts_Find_with_the_filters_text_the_cursor_at_its_end()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC")));
            Window window = Show(editor);

            editor.SearchText = "VIC";
            editor.OpenReplaceBar();
            Dispatcher.UIThread.RunJobs();

            TextBox find = editor.GetControl<TextBox>("ReplaceFindBox");

            Assert.Equal("VIC", find.Text);
            Assert.True(find.IsFocused);
            Assert.Equal(3, find.CaretIndex);
            Assert.Equal(find.SelectionStart, find.SelectionEnd);
            Assert.Equal("U2", Label(Assert.Single(editor.RowsShownForTests)));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** "FIND" FILTERS AND MARKS AS THE SEARCH BOX DOES, THE TEXT AS TYPED (owner request, 2026-11-07: "I want that
    // 'Find' works in same way as the filtering, where it will filter and highlight my search string (however, do not
    // allow the fancy search - only search for the string as-is and do not 'AND' multiple words or exclude with
    // minus)"). *** Two words are one phrase, a minus is a character - and while the bar is open the search box is
    // emptied and off ("do remove the possibility to also filter the table"), back once "Cancel" closes it, every row
    // shown again. No "On every sheet", and "Cancel" where "Close" was.
    // ###########################################################################################
    [Fact]
    public void Find_shows_only_the_rows_holding_the_text_as_typed_and_the_search_box_is_off_meanwhile()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "CPU-VIC")));
            Window window = Show(editor);

            editor.SearchText = "VIC";
            Assert.Equal(2, editor.RowsShownForTests.Count);

            editor.OpenReplaceBar();

            // The search box emptied and off; what it filtered by is "Find"'s now (owner request, 2026-11-07), applied as
            // typed - the same two rows.
            Assert.False(editor.GetControl<TextBox>("SearchBox").IsEnabled);
            Assert.Equal(string.Empty, editor.GetControl<TextBox>("SearchBox").Text);
            Assert.Equal("VIC", editor.GetControl<TextBox>("ReplaceFindBox").Text);
            Assert.Equal(2, editor.RowsShownForTests.Count);
            Assert.Null(editor.FindControl<CheckBox>("ReplaceEverySheetBox"));
            Assert.Null(editor.FindControl<Button>("CloseReplaceButton"));
            Assert.Equal("Cancel", editor.GetControl<Button>("CancelReplaceButton").Content);

            // The words as one phrase, not ANDed - "CPU VIC" is in no row; "CPU-VIC", minus and all, in one.
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "CPU VIC";
            editor.ApplyReplaceFindNowForTests();
            Assert.Empty(editor.RowsShownForTests);

            // A minus is a character: the row holding "-VIC", not every row without "VIC".
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "-VIC";
            editor.ApplyReplaceFindNowForTests();
            Assert.Equal("U3", Label(Assert.Single(editor.RowsShownForTests)));

            // The emptied search box's own TextChanged comes later (Avalonia posts it) and replaces nothing.
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("U3", Label(Assert.Single(editor.RowsShownForTests)));

            editor.GetControl<Button>("CancelReplaceButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.False(editor.IsReplaceBarOpen);
            Assert.True(editor.GetControl<TextBox>("SearchBox").IsEnabled);
            Assert.Equal(3, editor.RowsShownForTests.Count);

            window.Close();
        });
    }

    // ###########################################################################################
    // A table made read-only with the replace bar open - the Maintainer tab's, a submission's detail arriving after
    // Ctrl+H - lets go of the bar AND of its "Find" as the filter (code review, 2026-11-07: the bar went, the search box
    // came back empty, and the table went on showing only the rows "Find" held, with nothing saying why).
    // ###########################################################################################
    [Fact]
    public void A_table_made_read_only_lets_go_of_the_replace_bar_and_its_filter()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")));
            Window window = Show(editor);

            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "U1";
            editor.ApplyReplaceFindNowForTests();
            Assert.Equal("U1", Label(Assert.Single(editor.RowsShownForTests)));

            editor.IsReadOnly = true;

            Assert.False(editor.IsReplaceBarOpen);
            Assert.True(editor.GetControl<TextBox>("SearchBox").IsEnabled);
            Assert.Equal(["U1", "U2", "U3"], editor.RowsShownForTests.Select(Label));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** "COLUMNS": FIND AND REPLACE IN THE COLUMNS TICKED (owner request, 2026-11-07: "I need some possibility to
    // restrict replacement to specific columns, and not necessarily all"; a "Columns: All" drop-down, a tick per column
    // shown, "Yes, let us try this approach"). *** "CPU" is the friendly name of U1 and part of U2's technical name: with
    // the technical name unticked, Find shows U1 alone, marks nothing in that column, counts one, and Replace all
    // leaves U2 as it was. The choice stays with the sheet while another is shown, and "Every column" ticks all again.
    // ###########################################################################################
    [Fact]
    public void Columns_unticked_are_neither_found_marked_counted_nor_replaced_and_the_choice_stays_with_the_sheet()
    {
        UiTest.Run(() =>
        {
            var second = new ComponentEntry { BoardLabel = "U2", FriendlyName = "VIC", TechnicalNameOrValue = "CPU bus buffer" };
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), second, Component("U3", "SID")));
            Window window = Show(editor);
            int technical = Column(BoardWorkbookSchema.ColTechnicalNameOrValue);
            string technicalName = Components(editor).Columns[technical];

            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "CPU";
            editor.ApplyReplaceFindNowForTests();

            Assert.Equal("Columns: All", editor.ReplaceColumnsLabelForTests);
            Assert.Equal("2", editor.ReplaceCountForTests);
            Assert.Equal(["U1", "U2"], editor.RowsShownForTests.Select(Label));

            // The list as it opens: a tick per column shown, every one ticked.
            editor.GetControl<DropDownButton>("ReplaceColumnsButton").Flyout!.ShowAt(editor.GetControl<Control>("ReplaceColumnsButton"));
            Dispatcher.UIThread.RunJobs();
            List<CheckBox> ticks = editor.GetControl<StackPanel>("ReplaceColumnsPanel").Children.OfType<CheckBox>().ToList();
            Assert.Equal(Components(editor).Columns.Count, ticks.Count);
            Assert.All(ticks, tick => Assert.True(tick.IsChecked));

            ticks.Single(tick => (string?)tick.Content == technicalName).IsChecked = false;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal($"Columns: {ticks.Count - 1} of {ticks.Count}", editor.ReplaceColumnsLabelForTests);
            Assert.Equal("1", editor.ReplaceCountForTests);
            Assert.Equal(["U1"], editor.RowsShownForTests.Select(Label));

            // Nothing is marked in the column left out.
            BoardTableSearchMarks marks = BoardTableCellText.GetMarks(editor.GetControl<Control>("TableGrid"))!;
            Assert.False(marks.Search.Searches(Components(editor), technical));
            var cell = new BoardTableCellText { ColumnIndex = technical, DataContext = Row(editor, "U2") };
            BoardTableCellText.SetMarks(cell, marks);
            cell.CellText = "CPU bus buffer";
            Assert.Equal("CPU bus buffer", cell.Text);

            editor.GetControl<TextBox>("ReplaceWithBox").Text = "6510";
            editor.ReplaceAll();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("6510", Name(Row(editor, "U1")));
            Assert.Equal("CPU bus buffer", Row(editor, "U2").Cells[technical].Text);

            // Another sheet has every column; back on this one, its choice is as it was left.
            editor.SelectSheet(editor.DocumentForTests!.FindSheet(BoardWorkbookSchema.SheetBoardLinks)!);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("Columns: All", editor.ReplaceColumnsLabelForTests);

            editor.SelectSheet(Components(editor));
            Dispatcher.UIThread.RunJobs();
            Assert.Equal($"Columns: {ticks.Count - 1} of {ticks.Count}", editor.ReplaceColumnsLabelForTests);

            // "Unselect all", then the one column wanted (owner request, 2026-11-07: "so I can choose only one column").
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "bus";
            editor.ApplyReplaceFindNowForTests();
            editor.GetControl<Button>("ReplaceNoColumnButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Columns: None", editor.ReplaceColumnsLabelForTests);
            Assert.Equal("0", editor.ReplaceCountForTests);
            Assert.Empty(editor.RowsShownForTests);
            Assert.All(editor.GetControl<StackPanel>("ReplaceColumnsPanel").Children.OfType<CheckBox>(), tick => Assert.False(tick.IsChecked));

            editor.GetControl<StackPanel>("ReplaceColumnsPanel").Children.OfType<CheckBox>().Single(tick => (string?)tick.Content == technicalName).IsChecked = true;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal($"Columns: 1 of {ticks.Count}", editor.ReplaceColumnsLabelForTests);
            Assert.Equal("1", editor.ReplaceCountForTests);
            Assert.Equal(["U2"], editor.RowsShownForTests.Select(Label));

            editor.GetControl<Button>("ReplaceEveryColumnButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("Columns: All", editor.ReplaceColumnsLabelForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE TICKS UNDER "COLUMNS" ARE THE SHEET'S OWN, SO CHANGING THEM KEEPS THE TABLE ON IT (code review,
    // 2026-11-08). *** "Unselect all" leaves "Find" looking in no column of the sheet shown, so none of its rows show -
    // and the table moved to another sheet holding the text, the list still open now acting on that sheet: the one column
    // wanted could never be ticked while the text stood on any other sheet. And the list open as another sheet is shown
    // lists that sheet's columns, as its ticks then act on that sheet.
    // ###########################################################################################
    [Fact]
    public void Unticking_every_column_keeps_the_sheet_and_the_one_ticked_next_is_looked_in()
    {
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                Components =
                [
                    Component("U1", "CPU"),
                    new ComponentEntry { BoardLabel = "U2", FriendlyName = "VIC", TechnicalNameOrValue = "CPU bus buffer" },
                ],
                BoardLinks = [new BoardLinkEntry { Category = "Web", Name = "bus timing", Url = "https://example.com/bus" }],
            };

            BoardTableEditor editor = this.OpenEditor(board);
            Window window = Show(editor);
            int technical = Column(BoardWorkbookSchema.ColTechnicalNameOrValue);
            string technicalName = Components(editor).Columns[technical];
            StackPanel ticks = editor.GetControl<StackPanel>("ReplaceColumnsPanel");

            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "bus";
            editor.ApplyReplaceFindNowForTests();
            Assert.Same(Components(editor), editor.CurrentSheet);

            editor.GetControl<DropDownButton>("ReplaceColumnsButton").Flyout!.ShowAt(editor.GetControl<Control>("ReplaceColumnsButton"));
            Dispatcher.UIThread.RunJobs();
            editor.GetControl<Button>("ReplaceNoColumnButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Same(Components(editor), editor.CurrentSheet);
            Assert.Equal("Columns: None", editor.ReplaceColumnsLabelForTests);
            Assert.Empty(editor.RowsShownForTests);

            ticks.Children.OfType<CheckBox>().Single(tick => (string?)tick.Content == technicalName).IsChecked = true;
            Dispatcher.UIThread.RunJobs();

            Assert.Same(Components(editor), editor.CurrentSheet);
            Assert.Equal($"Columns: 1 of {Components(editor).Columns.Count}", editor.ReplaceColumnsLabelForTests);
            Assert.Equal(["U2"], editor.RowsShownForTests.Select(Label));
            Assert.Equal("1", editor.ReplaceCountForTests);

            // Another sheet shown with the list open: the list is that sheet's.
            BoardTableSheet links = editor.DocumentForTests!.FindSheet(BoardWorkbookSchema.SheetBoardLinks)!;
            editor.SelectSheet(links);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(links.Columns, ticks.Children.OfType<CheckBox>().Select(tick => (string?)tick.Content));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** "REPLACE ALL" PRESSED BEFORE "FIND" HAS CAUGHT UP REPLACES ON THE SHEET SHOWN, IN ITS TICKED COLUMNS (code
    // review, 2026-11-08). *** It looks for what "Find" holds first - and that moved the table to another sheet holding
    // the text when this one, in its ticked columns, held none, while the replace went on in the sheet it had started
    // on, now taken whole: a column unticked there was rewritten out of sight, the status line saying so under another
    // sheet.
    // ###########################################################################################
    [Fact]
    public void Replace_all_before_Find_has_caught_up_stays_on_the_sheet_and_its_ticked_columns()
    {
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", TechnicalNameOrValue = "VIC bus" }],
                BoardLinks = [new BoardLinkEntry { Category = "Web", Name = "VIC docs", Url = "https://example.com/vic" }],
            };

            BoardTableEditor editor = this.OpenEditor(board);
            Window window = Show(editor);
            int technical = Column(BoardWorkbookSchema.ColTechnicalNameOrValue);

            editor.OpenReplaceBar();
            editor.TickReplaceColumn(Components(editor).Columns[technical], ticked: false);

            // Typed and "Replace all" pressed at once - before the typing pause applies "Find".
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "VIC";
            editor.GetControl<TextBox>("ReplaceWithBox").Text = "SID";
            editor.ReplaceAll();

            Assert.Same(Components(editor), editor.CurrentSheet);
            Assert.Equal("VIC bus", Row(editor, "U1").Cells[technical].Text);
            BoardTableSheet links = editor.DocumentForTests!.FindSheet(BoardWorkbookSchema.SheetBoardLinks)!;
            Assert.Equal("VIC docs", links.Rows.Single().Cells[links.Columns.ToList().IndexOf("Name")].Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE COUNT ON "REPLACE ALL" IS WORKED OUT AGAIN ONLY WHEN IT CAN HAVE MOVED (code review, 2026-11-08). *** It
    // followed every toolbar update, and each count walks every cell ticked on the sheet - tens of thousands of string
    // operations at every arrow key on a sheet of 5,000 rows. The cursor moving changes nothing it counts; a cell
    // changed, or what "Replace with" holds, does.
    // ###########################################################################################
    [Fact]
    public void The_replace_count_is_worked_out_again_only_when_it_can_have_moved()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "CPU too"), Component("U3", "SID")));
            Window window = Show(editor);

            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "CPU";
            editor.ApplyReplaceFindNowForTests();
            Assert.Equal("2", editor.ReplaceCountForTests);
            int counted = editor.ReplaceCountsForTests;

            // The cursor moved about the rows "Find" shows: nothing it counts has moved.
            editor.SelectCell(Row(editor, "U1"), NameColumn);
            Dispatcher.UIThread.RunJobs();
            editor.SelectCell(Row(editor, "U2"), NameColumn);
            Dispatcher.UIThread.RunJobs();
            editor.SelectCell(Row(editor, "U2"), Column(BoardWorkbookSchema.ColBoardLabel));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(counted, editor.ReplaceCountsForTests);

            // A cell changed: counted again, and right.
            Row(editor, "U2").Cells[NameColumn].Text = "CPU and CPU";
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("3", editor.ReplaceCountForTests);
            Assert.True(editor.ReplaceCountsForTests > counted);

            // What "Replace with" holds: counted again - a place it would leave as it is counts none.
            counted = editor.ReplaceCountsForTests;
            editor.GetControl<TextBox>("ReplaceWithBox").Text = "CPU";
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("0", editor.ReplaceCountForTests);
            Assert.True(editor.ReplaceCountsForTests > counted);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** WHAT "REPLACE WITH" HOLDS IS DRAWN IN THE FOUND TEXT'S PLACE, AND "REPLACE ALL" COUNTS THE PLACES (owner
    // request, 2026-11-07: "it will overwrite the highlighted text with another color and then the content of this
    // text, so I can see directly what the new text will look like ... in the 'Replace all' button I would like to see
    // a counter badge for how many places it will replace"). *** The cells' marks carry the replacement, cut where a
    // replace cuts - "aaa" holds "aa" once - and the count is what Replace all then replaces.
    // ###########################################################################################
    [Fact]
    public void The_replacement_is_drawn_in_place_of_what_it_replaces_and_Replace_all_counts_the_places()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = this.OpenEditor(Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")));
            Window window = Show(editor);

            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "part";
            editor.ApplyReplaceFindNowForTests();

            BoardTableSearchMarks marks = BoardTableCellText.GetMarks(editor.GetControl<Control>("TableGrid"))!;
            Assert.True(marks.IsLiteral);
            Assert.Null(marks.Replacement);
            Assert.Equal("3", editor.ReplaceCountForTests);

            editor.GetControl<TextBox>("ReplaceWithBox").Text = "chip";
            Dispatcher.UIThread.RunJobs();

            marks = BoardTableCellText.GetMarks(editor.GetControl<Control>("TableGrid"))!;
            Assert.Equal("chip", marks.Replacement);
            Assert.NotNull(marks.ReplacementBackground);

            // A cell's text drawn as it will read: the replacement where the find stood.
            var cell = new BoardTableCellText();
            BoardTableCellText.SetMarks(cell, marks);
            cell.CellText = "part U1";
            Assert.Equal(["chip", " U1"], cell.Inlines!.OfType<Avalonia.Controls.Documents.Run>().Select(run => run.Text));
            Assert.Same(marks.ReplacementBackground, cell.Inlines!.OfType<Avalonia.Controls.Documents.Run>().First().Background);

            editor.GetControl<Button>("ReplaceAllButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Replaced [part] 3 times in 3 cells. Ctrl+Z undoes it.", Status(editor));
            Assert.Equal("0", editor.ReplaceCountForTests);

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

            Assert.False(editor.GetControl<Button>("ReplaceButton").IsVisible);

            Assert.False(editor.Paste("pasted"));
            Press(window, Key.H, RawInputModifiers.Control, PhysicalKey.H);

            Assert.Equal("VIC", Name(u2));
            Assert.False(editor.IsReplaceBarOpen);
            Assert.Equal("VIC", editor.CopySelectionText());

            window.Close();
        });
    }
}
