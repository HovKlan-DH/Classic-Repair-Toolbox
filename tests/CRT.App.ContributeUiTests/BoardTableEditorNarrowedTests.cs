using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// TYPING IN THE TABLE WHILE IT SHOWS ONLY SOME ROWS - a pill picked in the colour key, or a search
// (code review, 2026-11-02). Submit on a draft with errors opens the table under the Errors pill
// (TabDrafts.ShowErrorsBeforeSubmitAsync), so this is the main path for fixing errors.
//
// In a filtered view the grid's commit leaves it with NO current cell: the view takes the row out and
// puts it back. Tab read the cursor after that, found none, and FocusGrid put it on the sheet's first
// row - shown or not - in its first column: the next text typed overwrote that cell ("https://..." in
// the category of the first link, a component renamed under the Modified pill), or, the first row
// hidden, went nowhere. "Insert row" added at the end of the sheet, "Delete row" did nothing.
//
// Real keys in a shown window, on a draft's session - what ends up in the draft is what is asserted.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableEditorNarrowedTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private static readonly int CategoryColumn = BoardWorkbookSchema.BoardLinks.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColCategory);
    private static readonly int NameColumn = BoardWorkbookSchema.BoardLinks.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName);
    private static readonly int UrlColumn = BoardWorkbookSchema.BoardLinks.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColUrl);

    private static BoardLinkEntry Link(string category, string name = "", string url = "") =>
        new() { Category = category, Name = name, Url = url };

    private static BoardData Links(params BoardLinkEntry[] links) => new() { BoardLinks = [.. links] };

    private (BoardTableEditor Editor, DraftSession Session, Window Window) Open(BoardData draft, string sheet, BoardData? published = null)
    {
        string path = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        CachedWorkbooks.Write(path, draft);

        DraftSession session = DraftSession.Open(this.DraftsRoot, BoardKey)!;

        var editor = new BoardTableEditor { AutoSave = false };
        Assert.True(editor.Load(session, published));
        Assert.True(editor.ShowSheetPart(sheet, BoardTableSheetPart.Whole));

        var window = new Window { Content = editor, Width = 1300, Height = 600 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (editor, session, window);
    }

    private (BoardTableEditor Editor, DraftSession Session, Window Window) OpenLinks(params BoardLinkEntry[] links) =>
        this.Open(Links(links), BoardWorkbookSchema.SheetBoardLinks);

    private static BoardTableRow Row(BoardTableEditor editor, string category) =>
        editor.CurrentSheet!.Rows.Single(row => !row.IsDeleted && row.Cells[CategoryColumn].Text == category);

    private static DataGrid Grid(BoardTableEditor editor) => editor.GetControl<DataGrid>("TableGrid");

    private static List<BoardTableRow> Shown(BoardTableEditor editor) =>
        ((System.Collections.IEnumerable)Grid(editor).ItemsSource!).Cast<BoardTableRow>().ToList();

    private static void Put(BoardTableEditor editor, BoardTableRow row, int column)
    {
        editor.SelectCell(row, column);
        Grid(editor).Focus();
        Dispatcher.UIThread.RunJobs();
    }

    private static void Type(Window window, string text)
    {
        window.KeyTextInput(text);
        Dispatcher.UIThread.RunJobs();
    }

    private static void Press(Window window, Key key, RawInputModifiers modifiers = RawInputModifiers.None)
    {
        PhysicalKey physical = key switch
        {
            Key.Tab => PhysicalKey.Tab,
            Key.Enter => PhysicalKey.Enter,
            Key.F2 => PhysicalKey.F2,
            Key.Escape => PhysicalKey.Escape,
            Key.D => PhysicalKey.D,
            Key.F => PhysicalKey.F,
            Key.H => PhysicalKey.H,
            Key.Home => PhysicalKey.Home,
            Key.End => PhysicalKey.End,
            _ => PhysicalKey.None,
        };

        window.KeyPress(key, modifiers, physical, keySymbol: null);
        Dispatcher.UIThread.RunJobs();
    }

    // The links as the draft holds them: "category/name/url".
    private static List<string> InDraft(DraftSession session) =>
        session.Content.Board.BoardLinks.Select(link => $"{link.Category}/{link.Name}/{link.Url}").ToList();

    private static void Pick(BoardTableEditor editor, BoardTableRowKinds kind)
    {
        editor.TogglePill(kind);
        Dispatcher.UIThread.RunJobs();
    }

    // ###########################################################################################
    // Rows going together under a pill are looked at together (code review, 2026-11-03): each went over the
    // whole sheet again, so thousands selected on a long sheet held the UI thread for selected x sheet rows.
    // The document's own changes go over it too, so what is held is that the count does not grow with the rows.
    // ###########################################################################################
    [Fact]
    public void Rows_deleted_together_under_a_pill_do_not_go_over_the_sheet_per_row()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(Enumerable.Range(1, 40).Select(number => Link($"Category {number}")).ToArray());
            Pick(editor, BoardTableRowKinds.Errors);
            Assert.Equal(40, Shown(editor).Count);

            BoardTableSheet sheet = editor.CurrentSheet!;

            int TimesGoneOverDeleting(int rows)
            {
                int before = editor.RowsShownAskedForTests;

                sheet.DeleteRows(sheet.Rows.Take(rows).ToList(), out _);
                Dispatcher.UIThread.RunJobs();

                return editor.RowsShownAskedForTests - before;
            }

            int forTwo = TimesGoneOverDeleting(2);
            int forThirty = TimesGoneOverDeleting(30);

            Assert.Equal(forTwo, forThirty);
            Assert.Equal(sheet.Rows, Shown(editor));
            Assert.Equal(8, Shown(editor).Count);

            window.Close();
        });
    }

    // The finding, on each place a row can have among those shown: type the name, Tab, type the address.
    [Theory]
    [InlineData("Web")]
    [InlineData("Docs")]
    [InlineData("Tools")]
    public void Under_the_Errors_pill_Tab_goes_on_along_the_row_and_what_is_typed_next_lands_there(string category)
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"), Link("Tools"));
            Pick(editor, BoardTableRowKinds.Errors);
            Assert.Equal(3, Shown(editor).Count);

            Put(editor, Row(editor, category), NameColumn);
            Type(window, "Doc name");
            Press(window, Key.Tab);

            Assert.Same(Row(editor, category), editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);
            Assert.True(Grid(editor).IsKeyboardFocusWithin);

            Type(window, "https://example.org/");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                new[] { "Web", "Docs", "Tools" }.Select(each => each == category ? $"{each}/Doc name/https://example.org/" : $"{each}//"),
                InDraft(session));

            window.Close();
        });
    }

    [Theory]
    [InlineData("Web")]
    [InlineData("Docs")]
    [InlineData("Tools")]
    public void Under_the_Errors_pill_Shift_Tab_goes_back_along_the_row_and_what_is_typed_next_lands_there(string category)
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"), Link("Tools"));
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, category), UrlColumn);
            Type(window, "https://example.org/");
            Press(window, Key.Tab, RawInputModifiers.Shift);

            Assert.Same(Row(editor, category), editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "Doc name");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                new[] { "Web", "Docs", "Tools" }.Select(each => each == category ? $"{each}/Doc name/https://example.org/" : $"{each}//"),
                InDraft(session));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** HOME AND END STAY ON THE BOARD'S COLUMNS (owner report, 2026-11-07: up/down and home/end pressed, and the "Board
    // label" column looked wrong). *** The grid's own Home went to the narrow marker column - no cell of the board, so no
    // current cell at all - and left the table scrolled where End had taken it. With Ctrl, the first and the last row.
    // Under a pill, among the rows shown.
    // ###########################################################################################
    [Fact]
    public void Home_and_End_go_to_the_first_and_last_column_of_the_board_and_with_Ctrl_to_the_rows_at_either_end()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(
                Link("Web"), Link("Fine", "Fine", "https://example.org/"), Link("Docs"), Link("Tools"));
            int lastColumn = editor.CurrentSheet!.Columns.Count - 1;
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Docs"), NameColumn);

            Press(window, Key.End);
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(lastColumn, editor.CurrentCell?.ColumnIndex);

            Press(window, Key.Home);
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(0, editor.CurrentCell?.ColumnIndex);

            Press(window, Key.End, RawInputModifiers.Control);
            Assert.Same(Row(editor, "Tools"), editor.CurrentRow);
            Assert.Equal(lastColumn, editor.CurrentCell?.ColumnIndex);

            // The first row SHOWN: "Fine" has no error, so the pill hides it - "Web" is first anyway.
            Press(window, Key.Home, RawInputModifiers.Control);
            Assert.Same(Row(editor, "Web"), editor.CurrentRow);
            Assert.Equal(0, editor.CurrentCell?.ColumnIndex);
            Assert.True(Grid(editor).IsKeyboardFocusWithin);

            window.Close();
        });
    }

    // At the end of a row Tab wraps to the next row shown, past the rows the pill hides.
    [Fact]
    public void Under_a_pill_Tab_at_the_end_of_a_row_wraps_to_the_next_row_shown()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(
                Link("Web"), Link("Fine", "Fine", "https://example.org/"), Link("Docs"));
            int lastColumn = editor.CurrentSheet!.Columns.Count - 1;
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Web"), lastColumn);
            Press(window, Key.Tab);

            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(0, editor.CurrentCell!.ColumnIndex);

            Press(window, Key.Tab, RawInputModifiers.Shift);

            Assert.Same(Row(editor, "Web"), editor.CurrentRow);
            Assert.Equal(lastColumn, editor.CurrentCell!.ColumnIndex);

            window.Close();
        });
    }

    // The sheet's first row hidden by the pill: the cursor vanished, and the next typing went nowhere.
    [Fact]
    public void With_the_first_row_of_the_sheet_hidden_Tab_keeps_the_cursor_and_the_typing()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(
                Link("Fine", "Fine", "https://example.org/"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);
            Assert.Same(Row(editor, "Docs"), Assert.Single(Shown(editor)));

            Put(editor, Row(editor, "Docs"), NameColumn);
            Type(window, "Doc name");
            Press(window, Key.Tab);

            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "https://example.com/");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Fine/Fine/https://example.org/", "Docs/Doc name/https://example.com/"], InDraft(session));

            window.Close();
        });
    }

    // A paste straight after Tab lands on the cell Tab went to.
    [Fact]
    public void Under_a_pill_a_paste_after_Tab_lands_on_the_cell_Tab_went_to()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Docs"), NameColumn);
            Type(window, "Doc name");
            Press(window, Key.Tab);

            Assert.True(editor.Paste("https://example.org/"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web//", "Docs/Doc name/https://example.org/"], InDraft(session));

            window.Close();
        });
    }

    // The reviewer's other case: under the Modified pill the next value renamed U1.
    [Fact]
    public void Under_the_Modified_pill_Tab_does_not_send_the_next_value_to_the_first_rows_label()
    {
        UiTest.Run(() =>
        {
            ComponentEntry Component(string label, string name) => new() { BoardLabel = label, FriendlyName = name, TechnicalNameOrValue = $"part {label}", Category = "IC" };

            List<string> columns = BoardWorkbookSchema.Components.ColumnOrder.ToList();
            int label = columns.IndexOf(BoardWorkbookSchema.ColBoardLabel);
            int friendly = columns.IndexOf(BoardWorkbookSchema.ColFriendlyName);

            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                new BoardData { Components = [Component("U1", "CPU 6510"), Component("U2", "VIC 6569")] },
                BoardWorkbookSchema.SheetComponents,
                published: new BoardData { Components = [Component("U1", "CPU"), Component("U2", "VIC")] });

            Pick(editor, BoardTableRowKinds.Modified);

            BoardTableRow u1 = editor.CurrentSheet!.Rows.Single(row => row.Cells[label].Text == "U1");
            Put(editor, u1, friendly);
            Type(window, "typed1");
            Press(window, Key.Tab);

            Assert.Same(u1, editor.CurrentRow);
            Assert.Equal(friendly + 1, editor.CurrentCell!.ColumnIndex);

            Type(window, "typed2");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["U1", "U2"], session.Content.Board.Components.Select(component => component.BoardLabel));
            Assert.Equal("typed1", session.Content.Board.Components[0].FriendlyName);
            Assert.Equal("typed2", u1.Cells[friendly + 1].Text);

            window.Close();
        });
    }

    // A search's view is filtered too, though a row never leaves it for being edited (agreed case C).
    [Fact]
    public void Under_a_search_Tab_goes_on_along_the_row_and_what_is_typed_next_lands_there()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(
                Link("Other", "Other", "https://example.org/other"),
                Link("Web pick", "One"),
                Link("Docs pick", "Two"));

            editor.SearchText = "pick";
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(2, Shown(editor).Count);

            Put(editor, Row(editor, "Docs pick"), NameColumn);
            Type(window, "Doc name");
            Press(window, Key.Tab);

            Assert.Same(Row(editor, "Docs pick"), editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "https://example.org/");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                ["Other/Other/https://example.org/other", "Web pick/One/", "Docs pick/Doc name/https://example.org/"],
                InDraft(session));

            // Shift+Tab under the search too, on the first row shown.
            Put(editor, Row(editor, "Web pick"), UrlColumn);
            Type(window, "https://example.com/");
            Press(window, Key.Tab, RawInputModifiers.Shift);

            Assert.Same(Row(editor, "Web pick"), editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            window.Close();
        });
    }

    // A row's key and another cell changed: a new row and a red one where it was (BoardTableSheet) -
    // which arrive a moment after Tab has put the cursor on the next cell, and must leave it there.
    [Fact]
    public void Under_a_search_a_red_row_appearing_after_Tab_leaves_the_cursor_where_Tab_put_it()
    {
        UiTest.Run(() =>
        {
            BoardData links = Links(Link("Web pick", "One", "https://example.org/one"), Link("Docs pick", "Two", "https://example.org/two"));
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(links, BoardWorkbookSchema.SheetBoardLinks, published: links);

            editor.SearchText = "pick";
            Dispatcher.UIThread.RunJobs();

            BoardTableRow docs = Row(editor, "Docs pick");
            int region = editor.CurrentSheet!.Columns.ToList().IndexOf(BoardWorkbookSchema.ColRegion);

            Put(editor, docs, NameColumn);
            Type(window, "New");
            Press(window, Key.Tab);
            Type(window, "https://example.org/new");
            Press(window, Key.Tab);

            Assert.Contains(editor.CurrentSheet!.Rows, row => row.IsDeleted);
            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(region, editor.CurrentCell!.ColumnIndex);

            Type(window, "PAL");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            BoardLinkEntry saved = session.Content.Board.BoardLinks[1];
            Assert.Equal(["Docs pick", "New", "https://example.org/new", "PAL"], new[] { saved.Category, saved.Name, saved.Url, saved.Region });
            Assert.Equal("Web pick/One/https://example.org/one", InDraft(session)[0]);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A ROW THE CURSOR IS ON STAYS SHOWN UNTIL THE CURSOR LEAVES IT (owner's answer, 2026-11-02). ***
    // Its last error fixed, the Errors pill no longer shows it. It took the row out at once, and the
    // cursor had to land somewhere else - where the next text typed overwrote another row's cell.
    // What is typed next goes on along the row the person is working in, as long as they stay in it.
    // ###########################################################################################
    [Fact]
    public void A_row_whose_last_error_is_fixed_by_Tab_stays_until_the_cursor_leaves_it()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(
                Link("Web", "One"), Link("Docs", "Two"), Link("Tools", "Three"));
            int region = editor.CurrentSheet!.Columns.ToList().IndexOf(BoardWorkbookSchema.ColRegion);
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, web, UrlColumn);
            Type(window, "https://example.org/");
            Press(window, Key.Tab);

            Assert.Contains(web, Shown(editor));
            Assert.Same(web, editor.CurrentRow);
            Assert.Equal(region, editor.CurrentCell!.ColumnIndex);
            Assert.True(Grid(editor).IsKeyboardFocusWithin);

            Type(window, "PAL");
            Press(window, Key.Tab);

            // The link's note is its last column since 2026-11-06 (owner request): still the same row.
            Assert.Same(web, editor.CurrentRow);
            Assert.Equal(editor.CurrentSheet!.Columns.ToList().IndexOf(BoardWorkbookSchema.ColNote), editor.CurrentCell!.ColumnIndex);
            Assert.Contains(web, Shown(editor));

            Press(window, Key.Tab);

            // At the end of the row Tab goes on to the next row - the person moved - and only then is
            // the row they left no longer shown.
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(0, editor.CurrentCell!.ColumnIndex);
            Assert.DoesNotContain(web, Shown(editor));

            BoardLinkEntry saved = session.Content.Board.BoardLinks[0];
            Assert.Equal(["Web", "One", "https://example.org/", "PAL"], new[] { saved.Category, saved.Name, saved.Url, saved.Region });
            Assert.Equal(["Docs/Two/", "Tools/Three/"], InDraft(session).Skip(1));

            window.Close();
        });
    }

    [Fact]
    public void The_last_row_shown_stays_when_Tab_fixes_its_last_error_and_what_is_typed_next_lands_in_it()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"));
            int region = editor.CurrentSheet!.Columns.ToList().IndexOf(BoardWorkbookSchema.ColRegion);
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Docs"), UrlColumn);
            Type(window, "https://example.org/");
            Press(window, Key.Tab);

            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(region, editor.CurrentCell!.ColumnIndex);

            Type(window, "PAL");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/One/", "Docs/Two/https://example.org/"], InDraft(session));
            Assert.Equal("PAL", session.Content.Board.BoardLinks[1].Region);

            window.Close();
        });
    }

    [Fact]
    public void A_row_whose_last_error_is_fixed_by_Shift_Tab_stays_and_the_cursor_goes_back_along_it()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"), Link("Tools", "Three"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, UrlColumn);
            Type(window, "https://example.org/");
            Press(window, Key.Tab, RawInputModifiers.Shift);

            Assert.Contains(docs, Shown(editor));
            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "Renamed");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/One/", "Docs/Renamed/https://example.org/", "Tools/Three/"], InDraft(session));

            window.Close();
        });
    }

    // The very last error fixed: the row stays under the cursor, and nothing else can be overwritten.
    [Fact]
    public void The_only_row_shown_stays_when_Tab_fixes_its_last_error_and_what_is_typed_next_lands_in_it()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(
                Link("Fine", "Fine", "https://example.org/"), Link("Docs", "Two"));
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Docs"), UrlColumn);
            Type(window, "https://example.com/");
            Press(window, Key.Tab);

            Assert.Same(Row(editor, "Docs"), Assert.Single(Shown(editor)));
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);

            Type(window, "PAL");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Fine/Fine/https://example.org/", "Docs/Two/https://example.com/"], InDraft(session));
            Assert.Equal("PAL", session.Content.Board.BoardLinks[1].Region);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Insert row" pressed with a cell in its editor: the row goes below the row the cursor was on -
    // it went to the end of the sheet, the commit having taken the cursor away.
    // ###########################################################################################
    [Fact]
    public void Under_a_pill_Insert_row_with_a_cell_in_its_editor_adds_below_that_row()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"), Link("Tools"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, NameColumn);
            Type(window, "Doc name");

            editor.InsertRowBelow();
            Dispatcher.UIThread.RunJobs();

            BoardTableRow inserted = editor.CurrentSheet!.Rows[editor.CurrentSheet.Rows.IndexOf(docs) + 1];
            Assert.True(inserted.IsBlank);
            Assert.Same(inserted, editor.CurrentRow);
            Assert.Equal(["Web//", "Docs/Doc name/", "Tools//"], InDraft(session));

            window.Close();
        });
    }

    [Fact]
    public void Under_a_search_Insert_row_with_a_cell_in_its_editor_adds_below_that_row()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(Link("Web pick", "One"), Link("Docs pick", "Two"), Link("Tools pick", "Three"));
            editor.SearchText = "pick";
            Dispatcher.UIThread.RunJobs();

            BoardTableRow docs = Row(editor, "Docs pick");
            Put(editor, docs, UrlColumn);
            Type(window, "https://example.org/");

            editor.InsertRowBelow();
            Dispatcher.UIThread.RunJobs();

            BoardTableRow inserted = editor.CurrentSheet!.Rows[editor.CurrentSheet.Rows.IndexOf(docs) + 1];
            Assert.True(inserted.IsBlank);
            Assert.Same(inserted, editor.CurrentRow);

            window.Close();
        });
    }

    // "Delete row" pressed with a cell in its editor did nothing on that press.
    [Fact]
    public void Under_a_pill_Delete_row_with_a_cell_in_its_editor_deletes_on_that_press()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"), Link("Tools"));
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Docs"), NameColumn);
            Type(window, "Doc name");

            editor.DeleteRow();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web//", "Tools//"], InDraft(session));
            Assert.Same(Row(editor, "Tools"), editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            window.Close();
        });
    }

    [Fact]
    public void Under_a_search_Delete_row_with_a_cell_in_its_editor_deletes_on_that_press()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web pick", "One"), Link("Other", "Two"), Link("Docs pick", "Three"));
            editor.SearchText = "pick";
            Dispatcher.UIThread.RunJobs();

            Put(editor, Row(editor, "Web pick"), UrlColumn);
            Type(window, "https://example.org/");

            editor.DeleteRow();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Other/Two/", "Docs pick/Three/"], InDraft(session));
            Assert.Same(Row(editor, "Docs pick"), editor.CurrentRow);

            window.Close();
        });
    }

    // A paste fixing the row's last error leaves the row under the cursor, where the next Ctrl+V and
    // the next text typed go too.
    [Fact]
    public void A_paste_fixing_the_rows_last_error_leaves_the_row_under_the_cursor()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, web, UrlColumn);

            Assert.True(editor.Paste("https://example.org/"));
            Dispatcher.UIThread.RunJobs();

            Assert.Contains(web, Shown(editor));
            Assert.Same(web, editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);

            Assert.True(editor.Paste("https://example.org/again"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/One/https://example.org/again", "Docs/Two/"], InDraft(session));

            window.Close();
        });
    }

    // On a document too - the Maintainer tab's table, which saves when told and has its own undo.
    [Fact]
    public void In_a_document_under_a_pill_Tab_goes_on_along_the_row_and_what_is_typed_next_lands_there()
    {
        UiTest.Run(() =>
        {
            var editor = new BoardTableEditor();
            BoardTableDocument document = BoardTableDocument.Create(null, Links(Link("Web"), Link("Docs")));
            editor.Open(document);
            Assert.True(editor.ShowSheetPart(BoardWorkbookSchema.SheetBoardLinks, BoardTableSheetPart.Whole));

            var window = new Window { Content = editor, Width = 1300, Height = 600 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Docs"), NameColumn);
            Type(window, "Doc name");
            Press(window, Key.Tab);

            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "https://example.org/");
            Press(window, Key.Tab);

            // The row's last error fixed: it stays while the cursor is on it, and it is all there.
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(UrlColumn + 1, editor.CurrentCell!.ColumnIndex);
            Assert.Equal(
                ["Web//", "Docs/Doc name/https://example.org/"],
                document.ApplyTo(new BoardData()).BoardLinks.Select(link => $"{link.Category}/{link.Name}/{link.Url}"));

            window.Close();
        });
    }

    // Enter moves down under a pill, by the grid's own handling, and what is typed next lands there.
    [Fact]
    public void Under_a_pill_Enter_moves_down_and_what_is_typed_next_lands_there()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Web"), NameColumn);
            Type(window, "One");
            Press(window, Key.Enter);

            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "Two");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/One/", "Docs/Two/"], InDraft(session));

            window.Close();
        });
    }

    // Enter fixing a row's last error: the row leaves the view, the cursor stays on a row shown.
    [Fact]
    public void A_row_whose_last_error_is_fixed_by_Enter_leaves_the_view_and_the_cursor_stays_on_a_row_shown()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, web, UrlColumn);
            Type(window, "https://example.org/");
            Press(window, Key.Enter);

            Assert.DoesNotContain(web, Shown(editor));
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "https://example.com/");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/One/https://example.org/", "Docs/Two/https://example.com/"], InDraft(session));

            window.Close();
        });
    }

    // A click on another cell commits what was typed, and the cursor is on the cell clicked.
    [Fact]
    public void Under_a_pill_a_click_on_another_row_commits_and_what_is_typed_next_lands_there()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"), Link("Tools"));
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Web"), NameColumn);
            Type(window, "One");

            ClickCell(window, Row(editor, "Tools"), NameColumn);

            Assert.Same(Row(editor, "Tools"), editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "Three");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/One/", "Docs//", "Tools/Three/"], InDraft(session));

            window.Close();
        });
    }

    // The click after an edit that fixed the row's last error: it stays on the row clicked.
    [Fact]
    public void A_click_away_from_a_row_whose_last_error_it_fixed_keeps_the_cursor_on_the_row_clicked()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"), Link("Tools", "Three"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, web, UrlColumn);
            Type(window, "https://example.org/");

            ClickCell(window, Row(editor, "Tools"), NameColumn);

            Assert.DoesNotContain(web, Shown(editor));
            Assert.Same(Row(editor, "Tools"), editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);
            Assert.Equal(["Web/One/https://example.org/", "Docs/Two/", "Tools/Three/"], InDraft(session));

            window.Close();
        });
    }

    // Picking a pill with the sheet's first row hidden: the cursor goes to the first row SHOWN, or the
    // next text typed goes nowhere (FocusGrid put it on the hidden row, which cannot be current).
    [Fact]
    public void Focus_given_to_the_table_under_a_pill_puts_the_cursor_on_the_first_row_shown()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(
                Link("Fine", "Fine", "https://example.org/"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);

            editor.FocusGrid();
            Dispatcher.UIThread.RunJobs();

            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);

            // And on what the pill is there for: not over the row's category.
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "Doc name");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Fine/Fine/https://example.org/", "Docs/Doc name/"], InDraft(session));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE VIEW IS NEVER FILTERED, SO A COMMIT TAKES NOTHING AWAY (code review, 2026-11-02). *** A
    // filtered DataGridCollectionView takes the row out and puts it back at every commit, which left
    // the grid with no current cell - on Enter on the last row shown, Ctrl+D, Ctrl+F, Ctrl+H, anything
    // that commits. The grid shows a list of its own, and the view on it has no filter.
    // ###########################################################################################
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Under_a_pill_or_a_search_the_grid_is_given_a_view_with_no_filter(bool search)
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(Link("Web pick"), Link("Fine", "Fine", "https://example.org/"), Link("Docs pick"));

            if (search)
            {
                editor.SearchText = "pick";
            }
            else
            {
                Pick(editor, BoardTableRowKinds.Errors);
            }

            Dispatcher.UIThread.RunJobs();

            Assert.Equal(2, Shown(editor).Count);
            Assert.Null(((Avalonia.Collections.DataGridCollectionView)Grid(editor).ItemsSource!).Filter);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Enter on the last row shown

    [Fact]
    public void Enter_on_the_last_row_shown_keeps_the_cursor_there_and_what_is_typed_next_lands_there()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, NameColumn);
            Type(window, "Doc name");
            Press(window, Key.Enter);

            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);
            Assert.True(Grid(editor).IsKeyboardFocusWithin);

            // Tab goes on along the row, and what is typed next lands there.
            Press(window, Key.Tab);
            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "https://example.org/");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web//", "Docs/Doc name/https://example.org/"], InDraft(session));

            window.Close();
        });
    }

    [Fact]
    public void Enter_on_the_last_row_shown_under_a_search_keeps_the_cursor_and_Ctrl_V_is_not_refused()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(
                Link("Web pick", "One"), Link("Other", "Other"), Link("Docs pick", "Two"));
            editor.SearchText = "pick";
            Dispatcher.UIThread.RunJobs();

            BoardTableRow docs = Row(editor, "Docs pick");
            Put(editor, docs, NameColumn);
            Type(window, "Typed");
            Press(window, Key.Enter);

            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            Assert.True(editor.Paste("Pasted"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web pick/One/", "Other/Other/", "Docs pick/Pasted/"], InDraft(session));

            window.Close();
        });
    }

    [Fact]
    public void F2_and_Enter_on_the_last_row_shown_keep_the_cursor_there()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(Link("Web"), Link("Docs", "Doc"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, NameColumn);
            Press(window, Key.F2);
            Press(window, Key.Enter);

            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            window.Close();
        });
    }

    [Fact]
    public void In_a_document_Enter_on_the_last_row_shown_keeps_the_cursor_there()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, BoardTableDocument document, Window window) = OpenDocument(Links(Link("Web"), Link("Docs")), null, BoardWorkbookSchema.SheetBoardLinks);
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, NameColumn);
            Type(window, "Doc name");
            Press(window, Key.Enter);

            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            Press(window, Key.Tab);
            Type(window, "https://example.org/");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                ["Web//", "Docs/Doc name/https://example.org/"],
                document.ApplyTo(new BoardData()).BoardLinks.Select(link => $"{link.Category}/{link.Name}/{link.Url}"));

            window.Close();
        });
    }

    // The measured one: the Enter also fixed the row's last error, rows above remained, and the next
    // text overwrote the first row's Category ("https://w//w/").
    [Fact]
    public void Enter_fixing_the_last_rows_last_error_leaves_it_under_the_cursor_and_the_rows_above_alone()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, UrlColumn);
            Type(window, "https://example.org/");
            Press(window, Key.Enter);

            Assert.Contains(docs, Shown(editor));
            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);

            Press(window, Key.Tab);
            Type(window, "PAL");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/One/", "Docs/Two/https://example.org/"], InDraft(session));
            Assert.Equal("PAL", session.Content.Board.BoardLinks[1].Region);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Rows being filled in

    // ###########################################################################################
    // "Insert row" under a pill that does not show what the row is becoming - and the Maintainer tab
    // remembers a pick, so the pill may simply still be on. The first Tab took the new row out, the
    // cursor landed on the next row shown, and the next text overwrote it ("Docs/b2/..." became
    // "NewName/b2/...").
    // ###########################################################################################
    private static BoardData LinksBoth(string first, string second) =>
        Links(Link("Web", first, "https://example.org/web"), Link("Docs", second, "https://example.org/docs"));

    [Fact]
    public void Insert_row_under_the_Modified_pill_stays_shown_while_it_is_filled_in()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                LinksBoth("a2", "b2"), BoardWorkbookSchema.SheetBoardLinks, published: LinksBoth("a1", "b1"));
            Pick(editor, BoardTableRowKinds.Modified);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, web, NameColumn);
            editor.InsertRowBelow();
            Dispatcher.UIThread.RunJobs();

            BoardTableRow inserted = editor.CurrentRow!;
            Assert.True(inserted.IsBlank);
            Grid(editor).Focus();

            Type(window, "NewName");
            Press(window, Key.Tab);

            Assert.Same(inserted, editor.CurrentRow);
            Assert.Contains(inserted, Shown(editor));

            Type(window, "n2");
            Press(window, Key.Tab);
            Type(window, "https://example.org/new");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                ["Web/a2/https://example.org/web", "NewName/n2/https://example.org/new", "Docs/b2/https://example.org/docs"],
                InDraft(session));

            // It is added, not modified: once the cursor leaves it, the pill no longer shows it.
            ClickCell(window, Row(editor, "Docs"), NameColumn);

            Assert.DoesNotContain(inserted, Shown(editor));
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);

            window.Close();
        });
    }

    [Fact]
    public void In_a_document_Insert_row_under_the_Modified_pill_stays_shown_while_it_is_filled_in()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, BoardTableDocument document, Window window) = OpenDocument(
                LinksBoth("a2", "b2"), LinksBoth("a1", "b1"), BoardWorkbookSchema.SheetBoardLinks);
            Pick(editor, BoardTableRowKinds.Modified);

            Put(editor, Row(editor, "Web"), NameColumn);
            editor.InsertRowBelow();
            Dispatcher.UIThread.RunJobs();

            BoardTableRow inserted = editor.CurrentRow!;
            Grid(editor).Focus();

            Type(window, "NewName");
            Press(window, Key.Tab);

            Assert.Same(inserted, editor.CurrentRow);

            Type(window, "n2");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                ["Web/a2/https://example.org/web", "NewName/n2/", "Docs/b2/https://example.org/docs"],
                document.ApplyTo(new BoardData()).BoardLinks.Select(link => $"{link.Category}/{link.Name}/{link.Url}"));

            window.Close();
        });
    }

    // Under Warnings and Deleted too: on the Components sheet here, whose rows carry warnings of their
    // own, with a red row to land next to.
    [Theory]
    [InlineData(BoardTableRowKinds.Warnings)]
    [InlineData(BoardTableRowKinds.Deleted)]
    public void Insert_row_under_another_pill_stays_shown_while_it_is_filled_in(BoardTableRowKinds kind)
    {
        UiTest.Run(() =>
        {
            ComponentEntry Part(string label) => new() { BoardLabel = label, FriendlyName = $"name {label}", TechnicalNameOrValue = $"part {label}", Category = "IC" };

            List<string> columns = BoardWorkbookSchema.Components.ColumnOrder.ToList();
            int label = columns.IndexOf(BoardWorkbookSchema.ColBoardLabel);

            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                new BoardData { Components = [Part("U1"), Part("U3")] },
                BoardWorkbookSchema.SheetComponents,
                published: new BoardData { Components = [Part("U1"), Part("U2"), Part("U3")] });

            Pick(editor, kind);

            BoardTableRow shown = Shown(editor)[0];
            Put(editor, shown, label);
            editor.InsertRowBelow();
            Dispatcher.UIThread.RunJobs();

            BoardTableRow inserted = editor.CurrentRow!;
            Assert.True(inserted.IsBlank);
            Grid(editor).Focus();

            Type(window, "U9");
            Press(window, Key.Tab);

            Assert.Same(inserted, editor.CurrentRow);
            Assert.Contains(inserted, Shown(editor));

            Type(window, "typed");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("typed", session.Content.Board.Components.Single(component => component.BoardLabel == "U9").FriendlyName);
            Assert.Equal("name U1", session.Content.Board.Components.Single(component => component.BoardLabel == "U1").FriendlyName);
            Assert.Equal("name U3", session.Content.Board.Components.Single(component => component.BoardLabel == "U3").FriendlyName);

            window.Close();
        });
    }

    // A value typed back to the published one: the row is no longer Modified, and it stays until the
    // cursor leaves it - what is typed next does not go into the next row shown.
    [Fact]
    public void A_value_typed_back_to_the_published_one_stays_under_the_cursor_until_it_leaves()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                LinksBoth("a2", "b2"), BoardWorkbookSchema.SheetBoardLinks, published: LinksBoth("a1", "b2x"));
            Pick(editor, BoardTableRowKinds.Modified);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, web, NameColumn);
            Type(window, "a1");
            Press(window, Key.Tab);

            Assert.Contains(web, Shown(editor));
            Assert.Same(web, editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);

            Type(window, "https://example.org/typed");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/a1/https://example.org/typed", "Docs/b2/https://example.org/docs"], InDraft(session));

            window.Close();
        });
    }

    // A row with no cursor in it still goes, the moment it no longer is of a picked kind.
    [Fact]
    public void A_row_the_cursor_is_not_on_still_leaves_the_view_when_it_is_no_longer_of_a_picked_kind()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, Row(editor, "Docs"), NameColumn);

            web.Cells[UrlColumn].Text = "https://example.org/";
            Dispatcher.UIThread.RunJobs();

            Assert.DoesNotContain(web, Shown(editor));
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);

            window.Close();
        });
    }

    // Keeping a row is for what changes under the cursor: picking a pill hides the row the cursor is on
    // when that row is not of its kind, or the pick would not do what it says.
    [Fact]
    public void Picking_a_pill_or_typing_a_search_still_hides_the_row_the_cursor_is_on()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(Link("Fine", "Fine", "https://example.org/"), Link("Docs"));

            BoardTableRow fine = Row(editor, "Fine");
            Put(editor, fine, NameColumn);
            Pick(editor, BoardTableRowKinds.Errors);

            Assert.DoesNotContain(fine, Shown(editor));

            Pick(editor, BoardTableRowKinds.Errors);
            Assert.Contains(fine, Shown(editor));

            Put(editor, fine, NameColumn);
            editor.SearchText = "docs";
            Dispatcher.UIThread.RunJobs();

            Assert.DoesNotContain(fine, Shown(editor));

            window.Close();
        });
    }

    // ###########################################################################################
    // "Replace" takes the rows the table shows - and since 2026-11-07 its "Find" IS the filter (owner request: "I want
    // that 'Find' works in same way as the filtering"), so a find typed is a new search, which keeps no row on screen
    // for the cursor: the row fixed under the Errors pill leaves as the search box would let it go, and what the
    // table shows - the pill's rows holding the text - is what is replaced.
    // ###########################################################################################
    [Fact]
    public void Replace_all_reaches_the_rows_Find_shows_and_a_row_kept_for_the_cursor_goes_with_the_find()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "One"));
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Web"), UrlColumn);
            Type(window, "https://example.org/");
            Press(window, Key.Tab);

            Assert.Contains(Row(editor, "Web"), Shown(editor));

            editor.OpenReplaceBar();
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "One";
            editor.GetControl<TextBox>("ReplaceWithBox").Text = "Two";
            editor.ApplyReplaceFindNowForTests();

            Assert.DoesNotContain(Row(editor, "Web"), Shown(editor));
            Assert.Equal("1", editor.ReplaceCountForTests);

            editor.ReplaceAll();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web/One/https://example.org/", "Docs/Two/"], InDraft(session));

            window.Close();
        });
    }

    // ------------------------------------------------------------------ An undo on a sheet shown in two

    private static readonly int ImagePinColumn = BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColPin);
    private static readonly int ImageNameColumn = BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName);
    private static readonly int ImageFileColumn = BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFile);
    private static readonly int ImageTimeDivColumn = BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColTimeDiv);

    private static BoardData Images(params ComponentImageEntry[] images) => new()
    {
        Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
        ComponentImages = [.. images],
    };

    private static ComponentImageEntry Picture(string name) => new() { BoardLabel = "U8", Name = name, File = $"Component images/{name}.png" };

    private static ComponentImageEntry Baseline(string pin, string name) => new() { BoardLabel = "U8", Pin = pin, Name = name, File = $"Scope baseline/{name}.png" };

    private static BoardTableRow ImageRow(BoardTableEditor editor, string name) =>
        editor.CurrentSheet!.Rows.Single(row => !row.IsDeleted && row.Cells[ImageNameColumn].Text == name);

    private static void ShowPart(BoardTableEditor editor, BoardTableSheetPart part)
    {
        editor.SheetsPickedElsewhere = true;
        Assert.True(editor.ShowSheetPart(BoardWorkbookSchema.SheetComponentImages, part));
        Dispatcher.UIThread.RunJobs();
    }

    // ###########################################################################################
    // A board file holds the pictures first, then the baselines. So a picture added below the last picture
    // and undone leaves the FIRST BASELINE at its index. That row only slid there: the view stays, the cursor
    // on the nearest row it shows (code review, 2026-11-04: the view jumped to "Oscilloscope baseline", and
    // the left column's entry with it).
    // ###########################################################################################
    [Fact]
    public void An_undo_landing_on_a_row_of_the_other_part_leaves_the_view_where_it_is()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                Images(Picture("Top"), Baseline("1", "First"), Baseline("2", "Second")),
                BoardWorkbookSchema.SheetComponentImages);

            ShowPart(editor, BoardTableSheetPart.Pictures);

            Put(editor, ImageRow(editor, "Top"), ImageNameColumn);
            editor.InsertRowBelow();
            Dispatcher.UIThread.RunJobs();

            Put(editor, editor.CurrentRow!, ImageNameColumn);
            Type(window, "New");
            Press(window, Key.Tab);

            Assert.Same(ImageRow(editor, "New"), editor.CurrentSheet!.Rows[1]);

            editor.Undo();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(3, session.Content.Board.ComponentImages.Count);
            Assert.Equal(BoardTableSheetPart.Pictures, editor.SheetPart);
            Assert.Same(ImageRow(editor, "Top"), editor.CurrentRow);
            Assert.DoesNotContain(ImageRow(editor, "First"), Shown(editor));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE REPLACE BAR'S "FIND" FOLLOWS THE PART SHOWN, AND GOES WITH THE TABLE (code review, 2026-11-07). *** Its
    // columns are ticked per sheet and part, and "Find" took them once: with the bar open and "Oscilloscope baseline"
    // picked after "Pictures", the rows shown and marked went by the pictures' ticks while "Replace all" and its count went
    // by the baselines' - what was marked was not what would be replaced. And a table closed with the bar open left it
    // open, "Find" no longer the filter and the search box greyed out.
    // ###########################################################################################
    [Fact]
    public void The_replace_bars_Find_follows_the_part_shown_and_goes_with_the_table()
    {
        UiTest.Run(() =>
        {
            // The text in the baseline's name alone - not in its file's.
            (BoardTableEditor editor, _, Window window) = this.Open(
                Images(Picture("Top"), new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "zzz first", File = "Scope baseline/first.png" }),
                BoardWorkbookSchema.SheetComponentImages);

            ShowPart(editor, BoardTableSheetPart.Pictures);
            editor.OpenReplaceBar();

            // On the pictures, "Name" is left out of "Find".
            editor.TickReplaceColumn(editor.CurrentSheet!.Columns[ImageNameColumn], ticked: false);
            editor.GetControl<TextBox>("ReplaceFindBox").Text = "zzz";
            editor.ApplyReplaceFindNowForTests();
            Assert.Empty(editor.RowsShownForTests);

            // On the baselines every column is ticked: the name is looked in, shown, marked and counted alike.
            ShowPart(editor, BoardTableSheetPart.Baselines);

            Assert.Equal(BoardTableSheetPart.Baselines, editor.SheetPart);
            Assert.Equal("zzz first", Assert.Single(editor.RowsShownForTests).Cells[ImageNameColumn].Text);
            Assert.Equal("1", editor.ReplaceCountForTests);
            Assert.True(BoardTableCellText.GetMarks(Grid(editor))!.Search.Searches(editor.CurrentSheet!, ImageNameColumn));

            editor.Clear();

            Assert.False(editor.IsReplaceBarOpen);
            Assert.True(editor.GetControl<TextBox>("SearchBox").IsEnabled);

            window.Close();
        });
    }

    // A row the step really changed is what the undo goes to, whichever part shows.
    [Fact]
    public void An_undo_of_a_change_to_a_row_of_the_other_part_goes_to_that_part()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                Images(Picture("Top"), Baseline("1", "First")),
                BoardWorkbookSchema.SheetComponentImages);

            ShowPart(editor, BoardTableSheetPart.Pictures);

            Put(editor, ImageRow(editor, "Top"), ImageNameColumn);
            Type(window, "Renamed");
            Press(window, Key.Tab);

            Assert.Equal("Renamed", session.Content.Board.ComponentImages[0].Name);

            ShowPart(editor, BoardTableSheetPart.Baselines);

            editor.Undo();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Top", session.Content.Board.ComponentImages[0].Name);
            Assert.Equal(BoardTableSheetPart.Pictures, editor.SheetPart);
            Assert.Same(ImageRow(editor, "Top"), editor.CurrentRow);

            window.Close();
        });
    }

    // The step's own column may be one the part it goes to leaves out - the pin of a picture, here, which a
    // "Replace" on every sheet can reach and a key press cannot, the column being hidden. The cursor goes to a
    // column it has, or it lands nowhere and the next text typed is lost.
    [Fact]
    public void An_undo_going_to_a_part_without_the_steps_column_keeps_a_cursor()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                Images(Picture("Other"), Picture("Top"), Baseline("1", "First")),
                BoardWorkbookSchema.SheetComponentImages);

            ImageRow(editor, "Top").Cells[ImagePinColumn].Text = "5";
            editor.RefreshPendingForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("5", session.Content.Board.ComponentImages[1].Pin);

            ShowPart(editor, BoardTableSheetPart.Baselines);

            editor.Undo();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardTableSheetPart.Pictures, editor.SheetPart);
            Assert.Same(ImageRow(editor, "Top"), editor.CurrentRow);
            Assert.NotNull(editor.CurrentCell);
            Assert.Contains(editor.CurrentCell!.ColumnIndex, BoardTableSheetParts.ColumnsShown(BoardTableSheetPart.Pictures, editor.CurrentSheet!));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A ROW WORKED ON AMONG THE BASELINES STAYS THERE (code review, 2026-11-04). *** Left with no pin it
    // was a picture by then, and went to a table with no Pin column: "Insert row", a file typed, another row
    // clicked to look at a value, and only Ctrl+Z or deleting it got out. Here its pin can still be typed.
    // ###########################################################################################
    [Fact]
    public void A_new_row_left_with_no_pin_stays_among_the_baselines_and_can_be_given_one()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                Images(Picture("Top"), Baseline("1", "First"), Baseline("2", "Second")),
                BoardWorkbookSchema.SheetComponentImages);

            ShowPart(editor, BoardTableSheetPart.Baselines);

            Put(editor, ImageRow(editor, "First"), ImageNameColumn);
            editor.InsertRowBelow();
            Dispatcher.UIThread.RunJobs();

            BoardTableRow added = editor.CurrentRow!;

            Put(editor, added, ImageNameColumn);
            Type(window, "New");
            Press(window, Key.Tab);

            Put(editor, added, ImageFileColumn);
            Type(window, "Scope baseline/new.png");
            Press(window, Key.Tab);

            Put(editor, ImageRow(editor, "Second"), ImageNameColumn);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardTableSheetPart.Baselines, editor.SheetPart);
            Assert.Contains(added, Shown(editor));

            Put(editor, added, ImagePinColumn);
            Type(window, "7");
            Press(window, Key.Tab);

            Assert.Equal("7", added.Cells[ImagePinColumn].Text);
            Assert.Contains(added, Shown(editor));

            window.Close();
        });
    }

    // Its pin cleared, an existing baseline stays too - with its settings or, as here, without.
    [Fact]
    public void A_baseline_whose_pin_is_cleared_stays_among_the_baselines_and_shows_among_the_pictures_too()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                Images(Picture("Top"), Baseline("1", "First"), Baseline("2", "Second")),
                BoardWorkbookSchema.SheetComponentImages);

            ShowPart(editor, BoardTableSheetPart.Baselines);

            BoardTableRow row = ImageRow(editor, "Second");

            Put(editor, row, ImagePinColumn);
            row.Cells[ImagePinColumn].Text = string.Empty;
            Dispatcher.UIThread.RunJobs();

            Put(editor, ImageRow(editor, "First"), ImageNameColumn);
            Dispatcher.UIThread.RunJobs();

            Assert.Contains(row, Shown(editor));

            ShowPart(editor, BoardTableSheetPart.Pictures);

            Assert.Contains(row, Shown(editor));

            window.Close();
        });
    }

    // Only among the baselines: a row worked on among the pictures is a picture, and stays one.
    [Fact]
    public void A_picture_worked_on_among_the_pictures_does_not_show_among_the_baselines()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                Images(Picture("Top"), Baseline("1", "First")),
                BoardWorkbookSchema.SheetComponentImages);

            ShowPart(editor, BoardTableSheetPart.Pictures);

            Put(editor, ImageRow(editor, "Top"), ImageNameColumn);
            Type(window, "Renamed");
            Press(window, Key.Tab);

            BoardTableRow row = ImageRow(editor, "Renamed");

            ShowPart(editor, BoardTableSheetPart.Baselines);

            Assert.DoesNotContain(row, Shown(editor));

            window.Close();
        });
    }

    // ###########################################################################################
    // A pill on "Pictures" whose only answer is a problem in a column it hides (code review, 2026-11-04):
    // the row was listed with nothing marked, and the cursor, sent to a column the grid did not have, was
    // left nowhere - so what was typed next had no cell to go to. It goes to the row among the baselines.
    // ###########################################################################################
    [Fact]
    public void A_pill_answered_only_in_a_hidden_column_goes_to_the_baselines_with_the_cursor_on_the_problem()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(
                Images(
                    Picture("Top"),
                    new ComponentImageEntry { BoardLabel = "U8", Name = "Odd", File = "Component images/Odd.png", TimeDiv = "nonsense" }),
                BoardWorkbookSchema.SheetComponentImages);

            ShowPart(editor, BoardTableSheetPart.Pictures);

            Grid(editor).CurrentCell = DataGridCellInfo.Unset;
            editor.Filter = BoardTableRowKinds.Warnings;
            Dispatcher.UIThread.RunJobs();

            editor.FocusGrid();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardTableSheetPart.Baselines, editor.SheetPart);
            Assert.Same(ImageRow(editor, "Odd"), editor.CurrentRow);
            Assert.NotNull(editor.CurrentCell);
            Assert.Equal(ImageTimeDivColumn, editor.CurrentCell!.ColumnIndex);

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Paste with a row edit open

    // ###########################################################################################
    // Typed in a cell, another cell of the same row clicked: the row's edit is still open, and Ctrl+V
    // commits it. The second line of a block went into a row the pill hides, unseen ("Fine/keep me"
    // became "Fine/n2"), one value on two selected cells filled only the framed one, and the cursor
    // was gone so the next Ctrl+V was refused.
    // ###########################################################################################
    [Fact]
    public void A_block_pasted_after_a_click_in_the_same_row_goes_only_into_rows_shown()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(
                Link("Web"), Link("Fine", "keep me", "https://example.org/fine"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, web, UrlColumn);
            Type(window, "https://example.org/web");
            ClickCell(window, web, NameColumn);

            Assert.True(editor.Paste("n1\r\nn2"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                ["Web/n1/https://example.org/web", "Fine/keep me/https://example.org/fine", "Docs/n2/"],
                InDraft(session));

            // The cursor is still there: the next Ctrl+V is not refused, and lands where it was.
            Assert.Same(web, editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);
            Assert.True(editor.Paste("n3"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Web/n3/https://example.org/web", InDraft(session)[0]);

            window.Close();
        });
    }

    [Fact]
    public void A_block_pasted_under_a_search_after_a_click_in_the_same_row_goes_only_into_rows_shown()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(
                Link("Web pick", "One"), Link("Fine", "keep me"), Link("Docs pick", "Two"));
            editor.SearchText = "pick";
            Dispatcher.UIThread.RunJobs();

            BoardTableRow web = Row(editor, "Web pick");
            Put(editor, web, UrlColumn);
            Type(window, "https://example.org/web");
            ClickCell(window, web, NameColumn);

            Assert.True(editor.Paste("n1\r\nn2"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web pick/n1/https://example.org/web", "Fine/keep me/", "Docs pick/n2/"], InDraft(session));

            window.Close();
        });
    }

    [Fact]
    public void One_value_pasted_on_two_cells_selected_after_a_click_in_the_row_fills_both()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow web = Row(editor, "Web");
            Put(editor, web, CategoryColumn);
            Type(window, "Webby");
            ClickCell(window, web, NameColumn);
            ClickCell(window, web, UrlColumn, RawInputModifiers.Shift);

            Assert.Equal(2, editor.SelectedTableCells().Count);

            Assert.True(editor.Paste("both"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Webby/both/both", "Docs//"], InDraft(session));

            window.Close();
        });
    }

    // ------------------------------------------------------------------ Every other commit

    // ###########################################################################################
    // Every other way of committing under a pill: it left no cursor, or moved it to the first row shown.
    // ###########################################################################################
    [Fact]
    public void Ctrl_F_with_a_cell_in_its_editor_keeps_the_cursor_for_when_the_table_is_back()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(Link("Web"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, NameColumn);
            Type(window, "Doc name");

            Press(window, Key.F, RawInputModifiers.Control);

            Assert.True(editor.GetControl<TextBox>("SearchBox").IsFocused);
            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            window.Close();
        });
    }

    [Fact]
    public void Ctrl_H_with_a_cell_in_its_editor_and_the_bar_closed_again_keeps_the_cursor()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web"), Link("Docs"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, NameColumn);
            Type(window, "Doc name");

            Press(window, Key.H, RawInputModifiers.Control);
            Assert.True(editor.IsReplaceBarOpen);

            Press(window, Key.Escape);
            Assert.False(editor.IsReplaceBarOpen);

            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(NameColumn, editor.CurrentCell!.ColumnIndex);

            Press(window, Key.Tab);
            Type(window, "https://example.org/");
            Grid(editor).CommitEdit();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Web//", "Docs/Doc name/https://example.org/"], InDraft(session));

            window.Close();
        });
    }

    // ###########################################################################################
    // Another editor's change arriving with a cell in its editor: the table is built again from the
    // session, with new row objects. The cursor stays on the same row - also one the pill would not
    // show any more - and what is typed next lands there.
    // ###########################################################################################
    [Fact]
    public void A_change_from_another_editor_with_a_cell_in_its_editor_keeps_the_cursor_on_that_row()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"), Link("Tools", "Three"));
            Pick(editor, BoardTableRowKinds.Errors);

            Put(editor, Row(editor, "Docs"), UrlColumn);
            Type(window, "https://example.org/");

            session.Apply("Another editor", content => content with
            {
                Board = content.Board.WithRevisionDate("2026-11-02"),
            });
            Dispatcher.UIThread.RunJobs();

            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);
            Assert.Equal(UrlColumn, editor.CurrentCell!.ColumnIndex);
            Assert.Contains(Row(editor, "Docs"), Shown(editor));
            Assert.Equal(["Web/One/", "Docs/Two/https://example.org/", "Tools/Three/"], InDraft(session));

            window.Close();
        });
    }

    // The same when the pill no longer shows anything of the sheet worked on, and does on another: the
    // table is not to move there, or the next text would go into that sheet's row at the same place.
    [Fact]
    public void A_change_from_another_editor_stays_on_the_sheet_worked_on_when_a_pill_shows_nothing_of_it_any_more()
    {
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                BoardLinks = [new BoardLinkEntry { Category = "Docs", Name = "Two", Url = "" }],
                Credits = [new CreditEntry { Category = "Data", NameOrHandle = "" }],
            };

            (BoardTableEditor editor, DraftSession session, Window window) = this.Open(board, BoardWorkbookSchema.SheetBoardLinks);
            Pick(editor, BoardTableRowKinds.Errors);
            Assert.True(editor.ShowSheetPart(BoardWorkbookSchema.SheetBoardLinks, BoardTableSheetPart.Whole));

            Put(editor, Row(editor, "Docs"), UrlColumn);
            Type(window, "https://example.org/");

            session.Apply("Another editor", content => content with
            {
                Board = content.Board.WithRevisionDate("2026-11-02"),
            });
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardWorkbookSchema.SheetBoardLinks, editor.CurrentSheet!.Name);
            Assert.Same(Row(editor, "Docs"), editor.CurrentRow);

            window.Close();
        });
    }

    // The marker column of the row being worked in clicked: still that row, so it stays shown.
    [Fact]
    public void A_click_on_the_marker_column_of_the_row_kept_by_the_cursor_keeps_it_shown()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, _, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow docs = Row(editor, "Docs");
            Put(editor, docs, UrlColumn);
            Type(window, "https://example.org/");
            Press(window, Key.Tab);

            ClickCell(window, docs, -1);

            Assert.Contains(docs, Shown(editor));
            Assert.Same(docs, editor.CurrentRow);

            window.Close();
        });
    }

    // A row above the row in hand leaving the view must not take the row kept only by the cursor with it.
    [Fact]
    public void A_row_leaving_above_the_one_in_hand_leaves_that_one_in_hand()
    {
        UiTest.Run(() =>
        {
            (BoardTableEditor editor, DraftSession session, Window window) = this.OpenLinks(Link("Web", "One"), Link("Docs", "Two"), Link("Tools", "Three"));
            Pick(editor, BoardTableRowKinds.Errors);

            BoardTableRow web = Row(editor, "Web");
            BoardTableRow docs = Row(editor, "Docs");

            Put(editor, docs, UrlColumn);
            Type(window, "https://example.org/docs");
            Press(window, Key.Tab);
            Assert.Same(docs, editor.CurrentRow);

            web.Cells[UrlColumn].Text = "https://example.org/web";
            Dispatcher.UIThread.RunJobs();

            Assert.DoesNotContain(web, Shown(editor));
            Assert.Contains(docs, Shown(editor));
            Assert.Same(docs, editor.CurrentRow);
            Assert.Equal(UrlColumn + 1, editor.CurrentCell!.ColumnIndex);

            Dispatcher.UIThread.RunJobs();
            Assert.Contains(docs, Shown(editor));

            window.Close();
        });
    }

    private static (BoardTableEditor Editor, BoardTableDocument Document, Window Window) OpenDocument(BoardData draft, BoardData? published, string sheet)
    {
        var editor = new BoardTableEditor();
        BoardTableDocument document = BoardTableDocument.Create(published, draft);
        editor.Open(document);
        Assert.True(editor.ShowSheetPart(sheet, BoardTableSheetPart.Whole));

        var window = new Window { Content = editor, Width = 1300, Height = 600 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (editor, document, window);
    }

    private static void ClickControl(Window window, Control control)
    {
        Point point = control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    private static void ClickCell(Window window, BoardTableRow row, int column, RawInputModifiers modifiers = RawInputModifiers.None)
    {
        DataGridCell cell = window.GetVisualDescendants()
            .OfType<DataGridCell>()
            .Single(each => each.IsEffectivelyVisible && ReferenceEquals(each.DataContext, row) && each.OwningColumn?.Tag is int tag && tag == column);

        Point point = cell.TranslatePoint(new Point(cell.Bounds.Width / 2, cell.Bounds.Height / 2), window)!.Value;

        window.MouseDown(point, MouseButton.Left, modifiers);
        window.MouseUp(point, MouseButton.Left, modifiers);
        Dispatcher.UIThread.RunJobs();
    }
}
