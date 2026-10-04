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
// SELECTING SEVERAL ROWS AND DELETING THEM IN ONE GO (owner request, 2026-10-02: "how difficult
// would it be to mark multiple rows and then delete those in one go? E.g. ... I would like to mark
// all 8 pinout rows and delete those").
//
// Each test is one case the project owner agreed before the code was written, numbered as it was
// then - A and B were questions, answered "yes" (the button says how many) and "No" (the Delete key
// deletes no row). The rule itself - which rows go, one undo step, a component's rows on the other
// sheets - is BoardTableSheet.DeleteRows, pinned in CRT.Data.Tests; these pin the grid around it,
// with a real pointer: Shift+click and Ctrl+click the way a contributor selects.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableEditorSelectionTests
{
    private static readonly string[] Pinouts = ["CR13", "CR17", "CR20", "CR21", "CR22", "CR23", "CR102", "CR103"];

    private static ComponentImageEntry Image(string label, string name = "Pinout", string? file = null) =>
        new() { BoardLabel = label, Name = name, File = file ?? $"a/{label.ToLowerInvariant()}.png" };

    private static BoardData Board(IEnumerable<ComponentImageEntry> images) => new() { ComponentImages = [.. images] };

    // The owner's screenshot: eight Pinout rows, then R307 twice.
    private static List<ComponentImageEntry> Screenshot() =>
    [
        .. Pinouts.Select(label => Image(label)),
        Image("R307", "Pinout (secondary)", "x/resistor_3k3_5.png"),
        Image("R307", "Pinout (secondary)", "x/resistor.png")
    ];

    private static (Window Window, BoardTableEditor Editor) Shown(BoardData draft, BoardData? published)
    {
        var editor = new BoardTableEditor();
        BoardTableDocument document = BoardTableDocument.Create(published, draft);
        editor.Open(document);
        editor.SelectSheet(document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!);

        var window = new Window { Content = editor, Width = 1400, Height = 800 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (window, editor);
    }

    private static (Window Window, BoardTableEditor Editor) ShownUnchanged() =>
        Shown(Board(Screenshot()), Board(Screenshot()));

    private static int LabelColumn(BoardTableEditor editor) =>
        editor.CurrentSheet!.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel);

    private static string Label(BoardTableRow row) =>
        row.Cells[row.Sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel)].Text;

    private static BoardTableRow Row(BoardTableEditor editor, string label, bool deleted = false) =>
        editor.CurrentSheet!.Rows.First(row => row.IsDeleted == deleted && Label(row) == label);

    private static List<string> Selected(BoardTableEditor editor) => editor.SelectedRows().Select(Label).ToList();

    private static List<string> Deleted(BoardTableEditor editor) =>
        editor.CurrentSheet!.Rows.Where(row => row.IsDeleted).Select(Label).ToList();

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    private static DataGridCell CellOnScreen(Window window, BoardTableRow row, int columnIndex) =>
        window.GetVisualDescendants()
            .OfType<DataGridCell>()
            .Single(cell => cell.IsEffectivelyVisible && ReferenceEquals(cell.DataContext, row) && cell.OwningColumn?.Tag is int tag && tag == columnIndex);

    // A real click on the row's Board label cell - with Shift or Ctrl held when asked.
    private static void ClickRow(Window window, BoardTableEditor editor, BoardTableRow row, RawInputModifiers modifiers = RawInputModifiers.None)
    {
        Point point = CentreOf(window, CellOnScreen(window, row, LabelColumn(editor)));
        window.MouseDown(point, MouseButton.Left, modifiers);
        window.MouseUp(point, MouseButton.Left, modifiers);
        Dispatcher.UIThread.RunJobs();
    }

    private static void Click(Window window, Control control)
    {
        Point point = CentreOf(window, control);
        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    // CR13, then Shift+click CR103: the eight Pinout rows.
    private static void SelectThePinouts(Window window, BoardTableEditor editor)
    {
        ClickRow(window, editor, Row(editor, "CR13"));
        ClickRow(window, editor, Row(editor, "CR103"), RawInputModifiers.Shift);
    }

    private static Button DeleteButton(BoardTableEditor editor) => editor.GetControl<Button>("DeleteRowButton");

    // Case 1: Shift+click selects every row from the first click to it; Ctrl+click adds a row, or
    // takes it out again.
    [Fact]
    public void Shift_click_selects_every_row_between_and_Ctrl_click_adds_or_removes_one()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShownUnchanged();

            SelectThePinouts(window, editor);
            Assert.Equal(Pinouts, Selected(editor));

            ClickRow(window, editor, Row(editor, "R307"), RawInputModifiers.Control);
            Assert.Equal([.. Pinouts, "R307"], Selected(editor));

            ClickRow(window, editor, Row(editor, "CR17"), RawInputModifiers.Control);
            Assert.Equal([.. Pinouts.Where(label => label != "CR17"), "R307"], Selected(editor));

            window.Close();
        });
    }

    // ###########################################################################################
    // Case 1's other half: a plain click selects that one cell again - also on a cell INSIDE a
    // selection of several, and straight after a search has narrowed the rows under it. Otherwise
    // "Delete 8 rows" would stay armed after a click meant to pick one row.
    // ###########################################################################################
    [Fact]
    public void A_plain_click_inside_a_selection_of_several_selects_that_cell_alone()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShownUnchanged();

            SelectThePinouts(window, editor);
            Assert.Equal(Pinouts, Selected(editor));

            ClickRow(window, editor, Row(editor, "CR20"));

            Assert.Equal(["CR20"], Selected(editor));
            Assert.Same(Row(editor, "CR20"), editor.CurrentRow);
            Assert.Equal("Delete row", DeleteButton(editor).Content);

            // The same straight after a search has narrowed the rows under the selection - the case
            // the render showed.
            SelectThePinouts(window, editor);
            editor.SearchText = "pinout -cr2";
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            Dispatcher.UIThread.RunJobs();

            ClickRow(window, editor, Row(editor, "CR13"));

            Assert.Equal(["CR13"], Selected(editor));
            Assert.Same(Row(editor, "CR13"), editor.CurrentRow);

            window.Close();
        });
    }

    // ###########################################################################################
    // How several selected cells look (built 2026-10-02, judged by eye in a render): each carries
    // a translucent veil - the cell template's CurrencyVisual rectangle, filled with
    // BoardTable_Selection_Fill - over its OWN colour, which the cell keeps (a changed cell stays
    // orange underneath). One selected cell has no veil: its dashed frame says it.
    // ###########################################################################################
    [Fact]
    public void Several_selected_cells_are_veiled_over_their_own_colour_and_one_is_only_framed()
    {
        UiTest.Run(() =>
        {
            List<ComponentImageEntry> draft = Screenshot();
            draft[3] = Image("CR21", file: "b/cr21.png");
            (Window window, BoardTableEditor editor) = Shown(Board(draft), Board(Screenshot()));

            int file = editor.CurrentSheet!.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFile);
            DataGridCell changed() => CellOnScreen(window, Row(editor, "CR21"), file);
            Avalonia.Controls.Shapes.Rectangle Veil(DataGridCell cell) =>
                cell.GetVisualDescendants().OfType<Avalonia.Controls.Shapes.Rectangle>().First(part => part.Name == "CurrencyVisual");

            Avalonia.Media.Color? orange = (changed().Background as Avalonia.Media.ISolidColorBrush)?.Color;

            Application app = Application.Current!;
            Assert.True(app.TryGetResource("BoardTable_Selection_Fill", app.ActualThemeVariant, out object? fill));
            Avalonia.Media.Color veil = ((Avalonia.Media.ISolidColorBrush)fill!).Color;

            // One cell: framed (the rectangle shows its dashed stroke) but not veiled.
            ClickRow(window, editor, Row(editor, "CR13"));
            Assert.DoesNotContain(BoardTableEditor.SeveralSelectedClass, editor.GetControl<DataGrid>("TableGrid").Classes);
            Assert.NotEqual(veil, (Veil(CellOnScreen(window, Row(editor, "CR13"), LabelColumn(editor))).Fill as Avalonia.Media.ISolidColorBrush)?.Color);

            Point last = CentreOf(window, CellOnScreen(window, Row(editor, "CR103"), file));
            window.MouseDown(last, MouseButton.Left, RawInputModifiers.Shift);
            window.MouseUp(last, MouseButton.Left, RawInputModifiers.Shift);
            Dispatcher.UIThread.RunJobs();

            Assert.Contains(BoardTableEditor.SeveralSelectedClass, editor.GetControl<DataGrid>("TableGrid").Classes);
            Assert.True(Veil(changed()).IsVisible);
            Assert.Equal(veil, (Veil(changed()).Fill as Avalonia.Media.ISolidColorBrush)?.Color);
            Assert.Equal(orange, (changed().Background as Avalonia.Media.ISolidColorBrush)?.Color);

            // A cell outside the selection has none.
            Assert.False(Veil(CellOnScreen(window, Row(editor, "R307"), file)).IsVisible);

            window.Close();
        });
    }

    // Cases 2 and 3: "Delete row" deletes every selected row - here all published, so all eight stay
    // as red rows where they were - and ONE Ctrl+Z brings all eight back.
    [Fact]
    public void Delete_row_with_several_rows_selected_deletes_them_all_and_one_Ctrl_Z_brings_them_back()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShownUnchanged();

            SelectThePinouts(window, editor);
            Click(window, DeleteButton(editor));

            Assert.Equal(Pinouts, Deleted(editor));
            Assert.Equal(10, editor.CurrentSheet!.Rows.Count);
            Assert.Equal(2, editor.CurrentSheet.Rows.Count(row => !row.IsDeleted && Label(row) == "R307"));

            window.KeyPress(Key.Z, RawInputModifiers.Control, PhysicalKey.Z, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(Deleted(editor));

            window.Close();
        });
    }

    // Case 4: red rows in the selection are skipped - and with ONLY red rows selected there is
    // nothing to delete, so the button is greyed out.
    [Fact]
    public void With_only_red_rows_selected_Delete_row_is_greyed_out()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(
                Board(Screenshot().Skip(2)),
                Board(Screenshot()));

            ClickRow(window, editor, Row(editor, "CR13", deleted: true));
            ClickRow(window, editor, Row(editor, "CR17", deleted: true), RawInputModifiers.Shift);

            Assert.Equal(["CR13", "CR17"], Selected(editor));
            Assert.False(DeleteButton(editor).IsEnabled);

            // One live row among them: that one is what it deletes.
            ClickRow(window, editor, Row(editor, "CR20"), RawInputModifiers.Control);
            Assert.True(DeleteButton(editor).IsEnabled);

            Click(window, DeleteButton(editor));
            Assert.Equal(["CR13", "CR17", "CR20"], Deleted(editor));

            window.Close();
        });
    }

    // Case 6: with a pill picked, only the rows on screen can be selected - a Shift+click range
    // never takes in a row the filter hides, so a hidden row is never deleted.
    [Fact]
    public void With_a_pill_picked_a_Shift_click_range_never_reaches_a_hidden_row()
    {
        UiTest.Run(() =>
        {
            List<ComponentImageEntry> draft = Screenshot();
            foreach (string label in new[] { "CR13", "CR21", "CR103" })
            {
                int index = draft.FindIndex(image => image.BoardLabel == label);
                draft[index] = Image(label, file: $"b/{label}.png");
            }

            (Window window, BoardTableEditor editor) = Shown(Board(draft), Board(Screenshot()));
            editor.TogglePill(BoardTableRowKinds.Modified);
            Dispatcher.UIThread.RunJobs();

            SelectThePinouts(window, editor);
            Assert.Equal(["CR13", "CR21", "CR103"], Selected(editor));

            Click(window, DeleteButton(editor));
            Assert.Equal(["CR13", "CR21", "CR103"], Deleted(editor));

            window.Close();
        });
    }

    // Case 7: afterwards the cursor is where the first of the deleted rows was - on its red row
    // when it was published, on the row that moved up into its place when it was not.
    [Fact]
    public void After_deleting_several_rows_the_cursor_is_where_the_first_of_them_was()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(
                Board([Image("CR13"), Image("CR17"), Image("X1"), Image("X2"), Image("X3")]),
                Board([Image("CR13"), Image("CR17")]));

            // Added rows simply go: X3 moves up into X1's place.
            ClickRow(window, editor, Row(editor, "X1"));
            ClickRow(window, editor, Row(editor, "X2"), RawInputModifiers.Shift);
            Click(window, DeleteButton(editor));

            Assert.Equal(["CR13", "CR17", "X3"], editor.CurrentSheet!.Rows.Select(Label));
            Assert.Equal("X3", Label(editor.CurrentRow!));

            // Published rows stay as red rows: the cursor is on CR13's.
            ClickRow(window, editor, Row(editor, "CR13"));
            ClickRow(window, editor, Row(editor, "CR17"), RawInputModifiers.Shift);
            Click(window, DeleteButton(editor));

            Assert.Equal(["CR13", "CR17"], Deleted(editor));
            Assert.Same(Row(editor, "CR13", deleted: true), editor.CurrentRow);

            window.Close();
        });
    }

    // Case 8: dragging a row by its grip moves that row only - never the others selected.
    [Fact]
    public void Dragging_a_rows_grip_with_several_selected_moves_only_that_row()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShownUnchanged();

            ClickRow(window, editor, Row(editor, "CR13"));
            ClickRow(window, editor, Row(editor, "CR17"), RawInputModifiers.Shift);
            Assert.Equal(["CR13", "CR17"], Selected(editor));

            Avalonia.Controls.Primitives.DataGridRowHeader HeaderOf(BoardTableRow row) =>
                window.GetVisualDescendants().OfType<Avalonia.Controls.Primitives.DataGridRowHeader>()
                    .Single(header => ReferenceEquals(header.DataContext, row));

            Point from = CentreOf(window, HeaderOf(Row(editor, "CR23")));
            Point to = CentreOf(window, HeaderOf(Row(editor, "CR20")));

            window.MouseDown(from, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
            window.MouseMove(new Point(from.X, from.Y - 10), RawInputModifiers.LeftMouseButton);
            Dispatcher.UIThread.RunJobs();
            window.MouseMove(to, RawInputModifiers.LeftMouseButton);
            Dispatcher.UIThread.RunJobs();
            window.MouseUp(to, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                ["CR13", "CR17", "CR23", "CR20", "CR21", "CR22", "CR102", "CR103", "R307", "R307"],
                editor.CurrentSheet!.Rows.Select(Label));

            window.Close();
        });
    }

    // Case 9: copy and paste stay single-cell - with several rows selected they act on the cell with
    // the dashed frame, the one clicked last, and on nothing else.
    [Fact]
    public void Copy_and_paste_with_several_rows_selected_act_on_the_framed_cell_only()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShownUnchanged();

            ClickRow(window, editor, Row(editor, "CR13"));
            ClickRow(window, editor, Row(editor, "CR17"), RawInputModifiers.Shift);
            Assert.Equal(["CR13", "CR17"], Selected(editor));

            Assert.Equal("CR17", editor.CopyTextOfCurrentCell());

            Assert.True(editor.PasteIntoCurrentCell("CR99"));
            Assert.Equal("CR13", Label(editor.CurrentSheet!.Rows[0]));
            Assert.Equal("CR99", Label(editor.CurrentSheet.Rows[1]));

            window.Close();
        });
    }

    // Case A (answered "yes"): the button says how many rows it will delete when it is more than
    // one - counting only the rows it can delete.
    [Fact]
    public void Delete_row_says_how_many_rows_it_will_delete_when_several_are_selected()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShownUnchanged();

            ClickRow(window, editor, Row(editor, "CR13"));
            Assert.Equal("Delete row", DeleteButton(editor).Content);

            SelectThePinouts(window, editor);
            Assert.Equal("Delete 8 rows", DeleteButton(editor).Content);

            ClickRow(window, editor, Row(editor, "R307"));
            Assert.Equal("Delete row", DeleteButton(editor).Content);

            window.Close();
        });
    }

    // ###########################################################################################
    // Case B (answered "No"): the Delete key deletes no row - in Excel it empties cells, and a row
    // delete on a reflex key press is easy to miss.
    //
    // *** IT DID BEFORE THIS CHANGE, behind the table's back. *** The grid's CanUserDeleteRows is
    // on by default, so Delete on a selected cell took the row straight out of the sheet's list:
    // no red row for a published one, no undo step. Found while checking this case (2026-10-02).
    // ###########################################################################################
    [Fact]
    public void The_Delete_key_deletes_no_row_with_one_or_several_selected()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShownUnchanged();
            BoardTableHistory history = editor.CommitAndGetDocument()!.History;
            int steps = history.UndoCount;

            ClickRow(window, editor, Row(editor, "CR17"));
            window.KeyPress(Key.Delete, RawInputModifiers.None, PhysicalKey.Delete, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            SelectThePinouts(window, editor);
            window.KeyPress(Key.Delete, RawInputModifiers.None, PhysicalKey.Delete, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(10, editor.CurrentSheet!.Rows.Count);
            Assert.Empty(Deleted(editor));
            Assert.Equal(steps, history.UndoCount);

            window.Close();
        });
    }
}
