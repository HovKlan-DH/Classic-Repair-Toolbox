using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// SEEING ALL OF A LONG TEXT (owner requests, 2026-10-04) - BoardTableEditor.TextWrap.cs:
//
//   - every cell wraps its text within its column, and its row grows to fit ("remove that checkbox
//     again and then always apply wrap text on cells, if the column size gets reduced" - it was a
//     "Wrap text" check box earlier the same day);
//   - pressing a column heading frees the column cap a drag could not get past;
//   - double-clicking a heading's edge fits the column to the widest text in EVERY row, as Excel
//     does ("doubleclicking the column resizer should auto-fit column to show all text") - never
//     wider than the table can show. The grid's own double-click fit measured only the rows on
//     screen, and is off.
//
// Shown in a real window, so the rows' heights and the columns' widths are the grid's own layout.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableEditorTextWrapTests
{
    // Far wider than a column sizes itself to, as the owner's note on the SID was.
    private const string LongNote =
        "The SID (Sound Interface Device) is a mono sound chip with three voices, each with its own envelope and " +
        "waveform, and a filter shared between them; it is fed from the 9 V rail on the earlier boards and from " +
        "12 V on the later ones, so check which it is before replacing it.";

    // Wider than "Pinout", narrower than the table.
    private const string LongName = "Pinout, with the voltages read on a working board";

    private static BoardData Board() => new()
    {
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U5", Name = "Pinout", File = "a/6581.jpg", Note = LongNote },
            new ComponentImageEntry { BoardLabel = "U11", Name = "Pinout", File = "a/8721.jpg", Note = "Short." }
        ]
    };

    // Forty rows, the long name in the LAST - far below the window's bottom edge, so the grid has
    // built no cell for it.
    private static BoardData ManyRowsWithTheLongNameLast() => new()
    {
        ComponentImages =
        [
            .. Enumerable.Range(1, 40).Select(number => new ComponentImageEntry
            {
                BoardLabel = $"U{number}",
                Name = number == 40 ? LongName : "Pinout",
                File = $"a/{number}.jpg"
            })
        ]
    };

    // The editor on the board's Component images sheet, in document mode - no draft file needed.
    private static BoardTableEditor Editor(BoardData? board = null)
    {
        var editor = new BoardTableEditor();
        board ??= Board();

        editor.Open(BoardTableDocument.Create(board, board));
        editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponentImages)!);

        return editor;
    }

    private static Window Show(BoardTableEditor editor)
    {
        var window = new Window { Content = editor, Width = 1200, Height = 600 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    private static DataGrid Grid(BoardTableEditor editor) => editor.GetControl<DataGrid>("TableGrid");

    private static int ColumnOf(string name) => BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(name);

    private static int NoteColumn => ColumnOf(BoardWorkbookSchema.ColNote);

    private static int NameColumn => ColumnOf(BoardWorkbookSchema.ColName);

    private static BoardTableRow RowOf(BoardTableEditor editor, string label) =>
        editor.CurrentSheet!.Rows.Single(row => row.Cells[0].Text == label);

    private static DataGridCell Cell(Window window, BoardTableRow row, int column) =>
        window.GetVisualDescendants()
            .OfType<DataGridCell>()
            .Single(cell => cell.IsEffectivelyVisible && ReferenceEquals(cell.DataContext, row) && cell.OwningColumn?.Tag is int tag && tag == column);

    private static bool IsBuilt(Window window, BoardTableRow row) =>
        window.GetVisualDescendants().OfType<DataGridCell>().Any(cell => cell.IsEffectivelyVisible && ReferenceEquals(cell.DataContext, row));

    private static DataGridColumn Column(BoardTableEditor editor, int index) =>
        Grid(editor).Columns.Single(column => column.Tag is int tag && tag == index);

    private static DataGridColumnHeader Header(Window window, string name) =>
        window.GetVisualDescendants().OfType<DataGridColumnHeader>().First(header => Equals(header.Content, name));

    // Two quick presses on one spot are a double-click.
    private static void DoubleClick(Window window, Visual target, double x)
    {
        Point point = target.TranslatePoint(new Point(x, target.Bounds.Height / 2), window)!.Value;

        for (int click = 0; click < 2; click++)
        {
            window.MouseDown(point, MouseButton.Left);
            window.MouseUp(point, MouseButton.Left);
        }

        Dispatcher.UIThread.RunJobs();
    }

    // ###########################################################################################
    // From the start, with nothing to tick: the long note runs over several lines within its
    // column, and its row grows to fit it; a row with nothing long stays MinRowHeight tall.
    // ###########################################################################################
    [Fact]
    public void Every_cell_wraps_and_a_long_texts_row_grows_to_show_it_all()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = Editor();
            Window window = Show(editor);

            Assert.Null(editor.FindControl<CheckBox>("WrapTextCheckBox"));
            Assert.Contains("CellsWrap", Grid(editor).Classes);
            Assert.True(double.IsNaN(Grid(editor).RowHeight));

            // The grid's own double-click fit measures only the rows on screen - the table's own
            // fit replaces it (the double-click tests below).
            Assert.False(Grid(editor).CanUserResizeColumnsOnDoubleClick);

            DataGridCell longCell = Cell(window, RowOf(editor, "U5"), NoteColumn);
            DataGridCell shortCell = Cell(window, RowOf(editor, "U11"), NoteColumn);

            TextBlock text = longCell.GetVisualDescendants().OfType<TextBlock>().First(block => block.Text == LongNote);
            Assert.Equal(TextWrapping.Wrap, text.TextWrapping);
            Assert.True(longCell.Bounds.Height > 2 * shortCell.Bounds.Height, $"long [{longCell.Bounds.Height}] short [{shortCell.Bounds.Height}]");

            // A row with nothing long keeps the ordinary height - sized to its text alone it shrank
            // to one bare line, cramped against its neighbours (seen in a render).
            Assert.Equal(BoardTableEditor.MinRowHeight, shortCell.Bounds.Height, precision: 0);
            Assert.True(longCell.Bounds.Width <= BoardTableEditor.MaxColumnWidth + 1);

            window.Close();
        });
    }

    // "If the column size gets reduced": a narrower column wraps its text onto more lines.
    [Fact]
    public void A_column_made_narrower_wraps_its_text_onto_more_lines()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = Editor();
            Window window = Show(editor);

            double tallBefore = Cell(window, RowOf(editor, "U5"), NoteColumn).Bounds.Height;

            editor.FreeColumnWidths();
            Column(editor, NoteColumn).Width = new DataGridLength(200);
            Dispatcher.UIThread.RunJobs();

            double tallAfter = Cell(window, RowOf(editor, "U5"), NoteColumn).Bounds.Height;

            Assert.True(tallAfter > tallBefore, $"before [{tallBefore}] after [{tallAfter}]");
            Assert.Equal(BoardTableEditor.MinRowHeight, Cell(window, RowOf(editor, "U11"), NoteColumn).Bounds.Height, precision: 0);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** "I CANNOT EXPAND THE CELL" *** - every column sizes itself up to MaxColumnWidth, and the cap
    // held for a drag too. Pressing a column heading - where a drag starts - frees every column:
    // the width it has, no upper limit. Another sheet's columns are built capped again.
    // ###########################################################################################
    [Fact]
    public void Pressing_a_column_heading_lets_a_column_be_dragged_wider_than_it_sizes_itself()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = Editor();
            Window window = Show(editor);

            DataGridColumn note = Column(editor, NoteColumn);
            Assert.Equal(BoardTableEditor.MaxColumnWidth, note.MaxWidth);
            double widthBefore = note.ActualWidth;

            // The first heading - always on screen, where the Note column's may be past the right
            // edge. Pressing any heading frees them all.
            DataGridColumnHeader header = Header(window, BoardWorkbookSchema.ColBoardLabel);
            Point point = header.TranslatePoint(new Point(header.Bounds.Width / 2, header.Bounds.Height / 2), window)!.Value;

            window.MouseDown(point, MouseButton.Left);
            window.MouseUp(point, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.True(double.IsPositiveInfinity(note.MaxWidth));
            Assert.Equal(widthBefore, note.Width.Value, precision: 0);
            Assert.All(
                Grid(editor).Columns.Where(column => column.Tag is int tag && tag >= 0),
                column => Assert.True(double.IsPositiveInfinity(column.MaxWidth)));

            // Wider than the cap, as a drag makes it.
            note.Width = new DataGridLength(BoardTableEditor.MaxColumnWidth + 300);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(BoardTableEditor.MaxColumnWidth + 300, note.ActualWidth, precision: 0);

            editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!);
            Assert.All(
                Grid(editor).Columns.Where(column => column.Tag is int tag && tag >= 0),
                column => Assert.Equal(BoardTableEditor.MaxColumnWidth, column.MaxWidth));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** EVERY ROW, NOT THE ROWS ON SCREEN. *** A double-click on the Name heading's right edge - or
    // on the left edge of the heading after it, which is the same line - fits the column to the long
    // name in the last row, which the grid has not built: scrolled into view, it is one line. The
    // grid's own fit measured only the rows built, so the column stayed "Pinout" wide and the name
    // wrapped.
    // ###########################################################################################
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Double_clicking_a_headings_edge_fits_the_column_to_every_row_even_one_not_on_screen(bool fromTheNextHeadingsLeftEdge)
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = Editor(ManyRowsWithTheLongNameLast());
            Window window = Show(editor);

            BoardTableRow last = RowOf(editor, "U40");
            DataGridColumn name = Column(editor, NameColumn);
            double widthBefore = name.ActualWidth;

            Assert.False(IsBuilt(window, last));

            if (fromTheNextHeadingsLeftEdge)
            {
                DoubleClick(window, Header(window, BoardWorkbookSchema.ComponentImages.ColumnOrder.ElementAt(NameColumn + 1)), x: 2);
            }
            else
            {
                DataGridColumnHeader header = Header(window, BoardWorkbookSchema.ColName);
                DoubleClick(window, header, x: header.Bounds.Width - 2);
            }

            Assert.True(name.ActualWidth > widthBefore, $"before [{widthBefore}] after [{name.ActualWidth}]");

            Grid(editor).ScrollIntoView(last, name);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardTableEditor.MinRowHeight, Cell(window, last, NameColumn).Bounds.Height, precision: 0);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** NEVER WIDER THAN THE TABLE CAN SHOW. *** A note far wider than the window fits the table,
    // and wraps the rest - a column past the right edge would hide the end of every line. And the
    // column is then WHOLLY in view: it grows to the right, and left where it was its right half sat
    // past the edge (seen in a render) - just past the frozen marker column, out to the right edge.
    // ###########################################################################################
    [Fact]
    public void A_text_wider_than_the_table_fits_the_table_and_wraps_the_rest()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = Editor();
            Window window = Show(editor);

            DataGridColumn note = Column(editor, NoteColumn);
            DataGridColumnHeadersPresenter headers = window.GetVisualDescendants().OfType<DataGridColumnHeadersPresenter>().Single();
            double marker = Grid(editor).Columns.Single(column => column.Tag is -1).ActualWidth;

            editor.FitColumnToText(note);
            Dispatcher.UIThread.RunJobs();

            Assert.True(note.ActualWidth > BoardTableEditor.MaxColumnWidth, $"width [{note.ActualWidth}]");
            Assert.Equal(headers.Bounds.Width - marker, note.ActualWidth, precision: 0);
            Assert.True(Cell(window, RowOf(editor, "U5"), NoteColumn).Bounds.Height > BoardTableEditor.MinRowHeight);

            DataGridColumnHeader noteHeader = Header(window, BoardWorkbookSchema.ColNote);
            double left = noteHeader.TranslatePoint(default, headers)!.Value.X;
            Assert.InRange(left, marker - 2, marker + 2);
            Assert.InRange(left + noteHeader.Bounds.Width, headers.Bounds.Width - 2, headers.Bounds.Width + 2);

            window.Close();
        });
    }

    // The narrow marker column at the far left is never resized - a double-click on its edge (the
    // first data heading's left edge) changes nothing.
    [Fact]
    public void A_double_click_on_the_marker_columns_edge_fits_nothing()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = Editor();
            Window window = Show(editor);

            DataGridColumn marker = Grid(editor).Columns.Single(column => column.Tag is -1);
            DataGridColumn label = Column(editor, ColumnOf(BoardWorkbookSchema.ColBoardLabel));
            double markerBefore = marker.ActualWidth;
            double labelBefore = label.ActualWidth;

            DoubleClick(window, Header(window, BoardWorkbookSchema.ColBoardLabel), x: 2);

            Assert.Equal(markerBefore, marker.ActualWidth, precision: 0);
            Assert.Equal(labelBefore, label.ActualWidth, precision: 0);

            window.Close();
        });
    }
}
