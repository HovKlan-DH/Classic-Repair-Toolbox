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
using Handlers.Geometry;

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

    // ###########################################################################################
    // Owner request, 2026-11-11: "if the cell values in the column does not take more space than the header, then do
    // "auto-fit" column based on the header name. In some cases I see the header have 0px horizontal spacing to its
    // right-sided cell border". Every data column of short values - "Board label" holding "U5" and "U11" among them - is at
    // least as wide as its heading's name with the space before it repeated after it; and pressing the header row (a drag about
    // to start) lets it be made narrower again, as wide as it then is.
    // ###########################################################################################
    [Fact]
    public void No_column_starts_narrower_than_its_heading_with_space_after_the_name()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = Editor();
            Window window = Show(editor);
            Dispatcher.UIThread.RunJobs();

            List<DataGridColumnHeader> headers = window.GetVisualDescendants().OfType<DataGridColumnHeader>()
                .Where(header => header.OwningColumn?.Tag is int tag && tag >= 0 && header.Content is string)
                .ToList();

            Assert.NotEmpty(headers);

            foreach (DataGridColumnHeader header in headers)
            {
                string name = (string)header.Content!;
                TextBlock shown = header.GetVisualDescendants().OfType<TextBlock>().First(text => text.Text == name);
                double before = shown.TranslatePoint(default, header)!.Value.X;

                var typeface = new Typeface(shown.FontFamily, shown.FontStyle, shown.FontWeight, shown.FontStretch);
                using var layout = new Avalonia.Media.TextFormatting.TextLayout(name, typeface, shown.FontSize, foreground: null);
                double after = header.Bounds.Width - before - layout.WidthIncludingTrailingWhitespace;

                Assert.True(after >= before - 1, $"[{name}]: {before:0.0} before the name, {after:0.0} after it");
            }

            DataGridColumn label = Column(editor, ColumnOf(BoardWorkbookSchema.ColBoardLabel));
            Assert.True(label.MinWidth > BoardTableEditor.LeastColumnWidth, $"[{label.MinWidth}]");

            editor.FreeColumnWidths();
            Assert.Equal(BoardTableEditor.LeastColumnWidth, label.MinWidth);

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
                // The next heading shown: a picture's part of the sheet leaves some columns out.
                List<int> shown = BoardTableSheetParts.ColumnsShown(editor.SheetPart, editor.CurrentSheet!);

                DoubleClick(window, Header(window, editor.CurrentSheet!.Columns[shown[shown.IndexOf(NameColumn) + 1]]), x: 2);
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
    // The Board schematics sheet with nothing longer than its headings - "30%", "Red" - as the
    // owner's C64 250407 has it.
    private static BoardTableEditor SchematicsEditor()
    {
        var board = new BoardData
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Board layout", SchematicImageFile = "a/b.png", SchematicHighlightColor = "Red", SchematicHighlightOpacity = "30%" },
                new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "a/c.png", SchematicHighlightColor = "Red", SchematicHighlightOpacity = "30%" }
            ]
        };

        var editor = new BoardTableEditor();
        editor.Open(BoardTableDocument.Create(board, board));
        editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!);

        return editor;
    }

    private static DataGridColumn SchematicsColumn(BoardTableEditor editor, string name) =>
        Column(editor, BoardWorkbookSchema.BoardSchematics.ColumnOrder.ToList().IndexOf(name));

    // The text block a heading's name is drawn with.
    private static TextBlock HeadingText(DataGridColumnHeader header) =>
        header.GetVisualDescendants().OfType<TextBlock>().First(text => text.Text == (string)header.Content!);

    private static int Lines(TextBlock text) => text.TextLayout.TextLines.Count;

    private static double OneLineWidth(TextBlock text)
    {
        using var layout = new Avalonia.Media.TextFormatting.TextLayout(
            text.Text!, new Typeface(text.FontFamily, text.FontStyle, text.FontWeight, text.FontStretch), text.FontSize, foreground: null);
        return layout.WidthIncludingTrailingWhitespace;
    }

    // ###########################################################################################
    // *** A COLUMN FITTED TO ITS HEADING HAS AS MUCH ROOM AFTER THE NAME AS BEFORE IT (owner report,
    // 2026-10-09: double-clicking the edge after "Schematic highlight opacity" left a wide empty
    // band after the name). *** The fit measured the whole heading, whose template keeps room for
    // a sort arrow the table never shows; it now measures the name, with the heading's own space
    // before it repeated after it.
    // ###########################################################################################
    [Fact]
    public void A_column_fitted_to_its_heading_has_as_much_room_after_the_name_as_before_it()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = SchematicsEditor();
            Window window = Show(editor);

            DataGridColumn opacity = SchematicsColumn(editor, BoardWorkbookSchema.ColSchematicHighlightOpacity);
            editor.FitColumnToText(opacity);
            Dispatcher.UIThread.RunJobs();

            DataGridColumnHeader header = Header(window, BoardWorkbookSchema.ColSchematicHighlightOpacity);
            TextBlock name = HeadingText(header);
            double before = name.TranslatePoint(default, header)!.Value.X;
            double after = opacity.ActualWidth - before - OneLineWidth(name);

            // The same, give or take the slack that keeps the name off a second line and the
            // rounding to whole pixels.
            Assert.True(before > 0, $"before [{before}]");
            Assert.InRange(after, before - 1, before + ColumnAutoFitGeometry.RoundingSlack + 1);
            Assert.Equal(1, Lines(name));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A HEADING WRAPS LIKE ITS CELLS (owner request, 2026-10-09: "can the table header also
    // please wrap"). *** A column made narrower than its name shows the name on more lines, every
    // word of it, and the heading row grows to fit - it was cut off at one line.
    // ###########################################################################################
    [Fact]
    public void A_heading_narrower_than_its_name_wraps_onto_more_lines_and_the_heading_row_grows()
    {
        UiTest.Run(() =>
        {
            BoardTableEditor editor = SchematicsEditor();
            Window window = Show(editor);

            DataGridColumnHeader header = Header(window, BoardWorkbookSchema.ColSchematicHighlightOpacity);
            double oneLineHeight = HeadingText(header).Bounds.Height;
            double rowBefore = header.Bounds.Height;

            editor.FreeColumnWidths();
            SchematicsColumn(editor, BoardWorkbookSchema.ColSchematicHighlightOpacity).Width = new DataGridLength(90);
            Dispatcher.UIThread.RunJobs();

            TextBlock name = HeadingText(header);
            Assert.Equal(TextWrapping.Wrap, name.TextWrapping);
            Assert.True(name.Bounds.Height >= 2 * oneLineHeight - 1, $"name [{name.Bounds.Height}] against one line [{oneLineHeight}]");
            Assert.True(header.Bounds.Height > rowBefore, $"heading [{header.Bounds.Height}] against [{rowBefore}]");
            Assert.True(name.Bounds.Width <= header.Bounds.Width, $"name [{name.Bounds.Width}] in heading [{header.Bounds.Width}]");

            window.Close();
        });
    }

    // ###########################################################################################
    // Owner report, 2026-11-11: "I do not understand why the first row in table is higher than the other rows? I think it
    // should have same height as the other rows". Four rows with the same three-line "Purpose or note", its column scrolled
    // out of view: ProDataGrid measured every cell of the FIRST row whatever the scroll, so that row alone grew to the note.
    // Every row is as tall as the cells on screen - all of them short, then all of them tall once scrolled to the note, and
    // short again scrolled back.
    // ###########################################################################################
    [Fact]
    public void The_first_row_is_as_tall_as_the_cells_on_screen_as_every_other_row_is()
    {
        UiTest.Run(() =>
        {
            const string Note = "I have **__no__** idea what this does:\n- Nothing?\n- Something?";
            var board = new BoardData
            {
                Regions = ["PAL", "NTSC", "HEST", "EU"],
                Components =
                [
                    .. new[] { "PAL", "NTSC", "HEST", "EU" }.Select(region => new ComponentEntry
                    {
                        BoardLabel = "C1", TechnicalNameOrValue = "0.1 uF", Category = "Capacitor", Region = region, Package = "Disc", Purpose = Note,
                    }),
                ],
            };

            var editor = new BoardTableEditor();
            editor.Open(BoardTableDocument.Create(board, board));
            editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!);

            var window = new Window { Content = editor, Width = 1100, Height = 700 };
            window.Show();
            Dispatcher.UIThread.RunJobs();
            Dispatcher.UIThread.RunJobs();

            DataGrid grid = Grid(editor);
            List<double> Heights() => grid.GetVisualDescendants().OfType<DataGridRow>().OrderBy(row => row.Index).Select(row => Math.Round(row.Bounds.Height)).ToList();

            List<double> start = Heights();
            Assert.Equal(4, start.Count);
            // One line each - a row is its cells and its 1px grid line.
            Assert.Single(start.Distinct());
            Assert.True(start[0] < BoardTableEditor.MinRowHeight + 2, $"[{string.Join(", ", start)}]");

            grid.ScrollIntoView(null!, grid.Columns.Last());
            Dispatcher.UIThread.RunJobs();
            Dispatcher.UIThread.RunJobs();

            List<double> scrolled = Heights();
            Assert.Single(scrolled.Distinct());
            Assert.True(scrolled[0] > 2 * BoardTableEditor.MinRowHeight - 2, $"[{string.Join(", ", scrolled)}]");

            grid.ScrollIntoView(null!, grid.Columns.First(column => column.Tag is int tag && tag >= 0));
            Dispatcher.UIThread.RunJobs();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(start, Heights());

            window.Close();
        });
    }

    // ###########################################################################################
    // Code review, 2026-11-11: a row moved up to the top (Alt+Up on the second row) gets a container loaded at index 0 while
    // the grid moves the old first row's container down WITHOUT loading it again - and that one kept its off-screen cells held
    // to a plain row's height, so scrolled to the note it alone stayed one line tall, its note cut. Every row grows to the
    // note, the one that was first included.
    // ###########################################################################################
    [Fact]
    public void A_row_moved_above_the_first_leaves_the_old_first_row_free_to_grow()
    {
        UiTest.Run(() =>
        {
            const string Note = "I have **__no__** idea what this does:\n- Nothing?\n- Something?";
            var board = new BoardData
            {
                Regions = ["PAL", "NTSC", "HEST", "EU"],
                Components =
                [
                    .. new[] { "PAL", "NTSC", "HEST", "EU" }.Select(region => new ComponentEntry
                    {
                        BoardLabel = "C1", TechnicalNameOrValue = "0.1 uF", Category = "Capacitor", Region = region, Package = "Disc", Purpose = Note,
                    }),
                ],
            };

            var editor = new BoardTableEditor();
            editor.Open(BoardTableDocument.Create(board, board));
            editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!);

            var window = new Window { Content = editor, Width = 1100, Height = 700 };
            window.Show();
            Dispatcher.UIThread.RunJobs();
            Dispatcher.UIThread.RunJobs();

            DataGrid grid = Grid(editor);
            List<double> Heights() => grid.GetVisualDescendants().OfType<DataGridRow>().Where(row => row.IsVisible).OrderBy(row => row.Index).Select(row => Math.Round(row.Bounds.Height)).ToList();

            editor.SelectCell(editor.CurrentSheet!.Rows[1], 1);
            Dispatcher.UIThread.RunJobs();
            editor.MoveRowUp();
            Dispatcher.UIThread.RunJobs();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("NTSC", editor.CurrentSheet!.Rows[0].Cells[ColumnOfComponents(BoardWorkbookSchema.ColRegion)].Text);

            grid.ScrollIntoView(null!, grid.Columns.Last());
            Dispatcher.UIThread.RunJobs();
            Dispatcher.UIThread.RunJobs();

            List<double> scrolled = Heights();
            Assert.Equal(4, scrolled.Count);
            Assert.True(scrolled.All(height => height > 2 * BoardTableEditor.MinRowHeight - 2), $"[{string.Join(", ", scrolled)}]");

            window.Close();
        });
    }

    private static int ColumnOfComponents(string name) => BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(name);
}
