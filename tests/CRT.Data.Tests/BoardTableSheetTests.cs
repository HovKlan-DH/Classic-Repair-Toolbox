using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardTableSheet's block edits - a block pasted, one value pasted onto several cells, and fill
// down (CRT 4.0, 2026-10-25: the plan's step 0, "the table view gains block copy and paste, fill
// down, and find & replace"). The rest of the sheet is BoardTableDocumentTests' and
// BoardTableHistoryTests'.
//
// Each edit writes many cells, so each test also holds it to being ONE undo step: one Ctrl+Z must
// put back everything it did, and nothing else.
// ###########################################################################################
public sealed class BoardTableSheetTests
{
    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, TechnicalNameOrValue = $"part {label}" };

    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    private static BoardTableSheet Components(BoardTableDocument document) =>
        document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

    private static int Column(string column) =>
        BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(column);

    private static int LabelColumn => Column(BoardWorkbookSchema.ColBoardLabel);

    private static int NameColumn => Column(BoardWorkbookSchema.ColFriendlyName);

    private static string Label(BoardTableRow row) => row.Cells[LabelColumn].Text;

    private static string Name(BoardTableRow row) => row.Cells[NameColumn].Text;

    private static BoardTableRow Row(BoardTableSheet sheet, string label, bool deleted = false) =>
        sheet.Rows.Single(row => row.IsDeleted == deleted && Label(row) == label);

    // Every row as "label:name", ghosts marked "-", in order.
    private static List<string> Picture(BoardTableSheet sheet) =>
        sheet.Rows.Select(row => $"{(row.IsDeleted ? "-" : "")}{Label(row)}:{Name(row)}").ToList();

    private static BoardTableDocument Draft(params ComponentEntry[] components) =>
        BoardTableDocument.Create(null, Board(components));

    // ------------------------------------------------------------------ PasteBlock

    [Fact]
    public void A_block_lands_from_the_anchor_down_and_to_the_right()
    {
        BoardTableSheet sheet = Components(Draft(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")));

        BoardTablePasteResult result = sheet.PasteBlock(
            Row(sheet, "U1").Cells[NameColumn],
            [["CPU 6510", "MOS 6510"], ["VIC II", "MOS 6569"]],
            sheet.Rows,
            mayAddRows: true);

        Assert.Equal(["U1:CPU 6510", "U2:VIC II", "U3:SID"], Picture(sheet));
        Assert.Equal("MOS 6510", Row(sheet, "U1").Cells[NameColumn + 1].Text);
        Assert.Equal("MOS 6569", Row(sheet, "U2").Cells[NameColumn + 1].Text);

        Assert.Equal([Row(sheet, "U1"), Row(sheet, "U2")], result.Rows);
        Assert.Equal(NameColumn, result.FirstColumn);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(4, result.CellsChanged);
        Assert.Equal(0, result.RowsAdded);
    }

    // A red row is not in the draft and cannot be typed into: the block goes on into the next row.
    [Fact]
    public void A_red_row_in_the_way_is_passed_over()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")),
            Board(Component("U1", "CPU"), Component("U3", "SID")));
        BoardTableSheet sheet = Components(document);

        Assert.Equal(["U1:CPU", "-U2:VIC", "U3:SID"], Picture(sheet));

        sheet.PasteBlock(Row(sheet, "U1").Cells[NameColumn], [["one"], ["two"]], sheet.Rows, mayAddRows: true);

        Assert.Equal(["U1:one", "-U2:VIC", "U3:two"], Picture(sheet));
    }

    // Rows past the end are added as "Insert row" adds them - and so placed in their category on
    // save, like any row added in this sitting.
    [Fact]
    public void Rows_past_the_end_are_added_at_the_end_of_the_sheet()
    {
        BoardTableDocument document = Draft(Component("U1", "CPU"));
        BoardTableSheet sheet = Components(document);

        BoardTablePasteResult result = sheet.PasteBlock(
            Row(sheet, "U1").Cells[LabelColumn],
            [["U1", "CPU"], ["U8", "PLA"], ["U9", "ROM"]],
            sheet.Rows,
            mayAddRows: true);

        Assert.Equal(["U1:CPU", "U8:PLA", "U9:ROM"], Picture(sheet));
        Assert.Equal(2, result.RowsAdded);
        Assert.Equal(3, result.Rows.Count);
        Assert.All(sheet.Rows.Skip(1), row => Assert.True(row.IsInsertedThisSession));
        Assert.True(document.HasUnsavedChanges);
    }

    // Narrowed by a pill or a search, a row added at the end would be out of sight - so none is.
    [Fact]
    public void While_rows_may_not_be_added_the_rest_of_the_block_is_left_out()
    {
        BoardTableSheet sheet = Components(Draft(Component("U1", "CPU"), Component("U2", "VIC")));

        BoardTablePasteResult result = sheet.PasteBlock(
            Row(sheet, "U1").Cells[NameColumn],
            [["one"], ["two"], ["three"]],
            [Row(sheet, "U1")],
            mayAddRows: false);

        Assert.Equal(["U1:one", "U2:VIC"], Picture(sheet));
        Assert.Equal(2, result.RowsLeftOut);
        Assert.Equal(0, result.RowsAdded);
    }

    [Fact]
    public void Cells_past_the_last_column_are_left_out()
    {
        BoardTableSheet sheet = Components(Draft(Component("U1", "CPU")));
        int last = sheet.Columns.Count - 1;

        BoardTablePasteResult result = sheet.PasteBlock(Row(sheet, "U1").Cells[last], [["described", "x", "y"]], sheet.Rows, mayAddRows: true);

        Assert.Equal("described", Row(sheet, "U1").Cells[last].Text);
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal(2, result.ColumnsLeftOut);
    }

    // Copied out of a spreadsheet with its heading row: the headings are never data.
    [Fact]
    public void A_first_row_of_the_sheets_own_column_headings_is_left_out()
    {
        BoardTableSheet sheet = Components(Draft(Component("U1", "CPU")));

        BoardTablePasteResult result = sheet.PasteBlock(
            Row(sheet, "U1").Cells[LabelColumn],
            [[BoardWorkbookSchema.ColBoardLabel, BoardWorkbookSchema.ColFriendlyName.ToUpperInvariant()], ["U1", "CPU 6510"]],
            sheet.Rows,
            mayAddRows: true);

        Assert.True(result.HeadingsLeftOut);
        Assert.Equal(["U1:CPU 6510"], Picture(sheet));
    }

    // Only a block's FIRST row, only where it names the columns it would land in.
    [Fact]
    public void Headings_landing_in_other_columns_are_data()
    {
        BoardTableSheet sheet = Components(Draft(Component("U1", "CPU"), Component("U2", "VIC")));

        BoardTablePasteResult result = sheet.PasteBlock(
            Row(sheet, "U1").Cells[NameColumn],
            [[BoardWorkbookSchema.ColBoardLabel], ["two"]],
            sheet.Rows,
            mayAddRows: true);

        Assert.False(result.HeadingsLeftOut);
        Assert.Equal($"U1:{BoardWorkbookSchema.ColBoardLabel}", Picture(sheet)[0]);
    }

    [Fact]
    public void A_shorter_block_row_writes_only_the_cells_it_has()
    {
        BoardTableSheet sheet = Components(Draft(Component("U1", "CPU"), Component("U2", "VIC")));

        sheet.PasteBlock(Row(sheet, "U1").Cells[NameColumn], [["one", "first part"], ["two"]], sheet.Rows, mayAddRows: true);

        Assert.Equal("first part", Row(sheet, "U1").Cells[NameColumn + 1].Text);
        Assert.Equal("part U2", Row(sheet, "U2").Cells[NameColumn + 1].Text);
    }

    [Fact]
    public void A_whole_paste_is_one_undo_step_added_rows_and_all()
    {
        BoardTableDocument document = Draft(Component("U1", "CPU"), Component("U2", "VIC"));
        BoardTableSheet sheet = Components(document);
        List<string> before = Picture(sheet);

        sheet.PasteBlock(Row(sheet, "U2").Cells[LabelColumn], [["U2", "VIC II"], ["U8", "PLA"], ["U9", "ROM"]], sheet.Rows, mayAddRows: true);
        Assert.Equal(1, document.History.UndoCount);

        BoardTableHistoryResult? undone = document.History.Undo();

        Assert.Equal(before, Picture(sheet));
        Assert.False(document.HasUnsavedChanges);

        // Ctrl+Z puts the cursor on the cell the block landed on, not on a row added at the end.
        Assert.Same(Row(sheet, "U2"), undone!.Row);

        document.History.Redo();
        Assert.Equal(["U1:CPU", "U2:VIC II", "U8:PLA", "U9:ROM"], Picture(sheet));
    }

    [Fact]
    public void Nothing_is_pasted_onto_a_red_row()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1", "CPU"), Component("U2", "VIC")), Board(Component("U1", "CPU")));
        BoardTableSheet sheet = Components(document);

        BoardTablePasteResult result = sheet.PasteBlock(Row(sheet, "U2", deleted: true).Cells[NameColumn], [["one"], ["two"]], sheet.Rows, mayAddRows: true);

        Assert.Same(BoardTablePasteResult.Nothing, result);
        Assert.False(document.History.CanUndo);
        Assert.Equal(["U1:CPU", "-U2:VIC"], Picture(sheet));
    }

    // The colours follow at once: the sheet is refreshed by the paste itself, not left for the UI.
    [Fact]
    public void A_pasted_block_is_coloured_as_soon_as_it_is_in()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1", "CPU")), Board(Component("U1", "CPU")));
        BoardTableSheet sheet = Components(document);

        sheet.PasteBlock(Row(sheet, "U1").Cells[NameColumn], [["CPU 6510"]], sheet.Rows, mayAddRows: true);

        Assert.Equal(BoardTableCellState.Modified, Row(sheet, "U1").Cells[NameColumn].State);
        Assert.Equal(1, sheet.ChangeCount);
        Assert.False(sheet.NeedsRefresh);
    }

    // ------------------------------------------------------------------ PasteIntoCells

    [Fact]
    public void One_value_fills_every_selected_cell_passing_red_rows_over_as_one_undo_step()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")),
            Board(Component("U1", "CPU"), Component("U3", "SID")));
        BoardTableSheet sheet = Components(document);
        BoardTableRow ghost = Row(sheet, "U2", deleted: true);

        int changed = sheet.PasteIntoCells(
            [Row(sheet, "U1").Cells[NameColumn], ghost.Cells[NameColumn], Row(sheet, "U3").Cells[NameColumn]],
            "chip");

        Assert.Equal(2, changed);
        Assert.Equal(["U1:chip", "-U2:VIC", "U3:chip"], Picture(sheet));

        document.History.Undo();
        Assert.Equal(["U1:CPU", "-U2:VIC", "U3:SID"], Picture(sheet));
        Assert.False(document.History.CanUndo);
    }

    // ------------------------------------------------------------------ FillDown

    [Fact]
    public void The_top_cell_of_each_column_fills_the_cells_below_it_as_one_undo_step()
    {
        BoardTableDocument document = Draft(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID"));
        BoardTableSheet sheet = Components(document);
        int part = Column(BoardWorkbookSchema.ColTechnicalNameOrValue);

        // Two columns: the name over three rows, the part over the first two.
        int changed = sheet.FillDown(
        [
            Row(sheet, "U3").Cells[NameColumn],
            Row(sheet, "U1").Cells[NameColumn],
            Row(sheet, "U2").Cells[NameColumn],
            Row(sheet, "U1").Cells[part],
            Row(sheet, "U2").Cells[part],
        ]);

        Assert.Equal(3, changed);
        Assert.Equal(["U1:CPU", "U2:CPU", "U3:CPU"], Picture(sheet));
        Assert.Equal("part U1", Row(sheet, "U2").Cells[part].Text);
        Assert.Equal("part U3", Row(sheet, "U3").Cells[part].Text);

        document.History.Undo();
        Assert.Equal(["U1:CPU", "U2:VIC", "U3:SID"], Picture(sheet));
        Assert.False(document.History.CanUndo);
    }

    // A red row's text is what is published, and shown - it may be filled FROM, never into.
    [Fact]
    public void A_red_row_can_be_filled_from_but_is_never_filled()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")),
            Board(Component("U2", "VIC"), Component("U3", "SID")));
        BoardTableSheet sheet = Components(document);

        Assert.Equal(["-U1:CPU", "U2:VIC", "U3:SID"], Picture(sheet));

        sheet.FillDown([Row(sheet, "U1", deleted: true).Cells[NameColumn], Row(sheet, "U2").Cells[NameColumn]]);
        Assert.Equal(["-U1:CPU", "U2:CPU", "U3:SID"], Picture(sheet));
    }

    // ###########################################################################################
    // Told apart from "the cells already hold it" (code review, 2026-10-25): every cell the paste or
    // the fill would write is on a red row - and it is said as that.
    // ###########################################################################################
    [Fact]
    public void A_paste_or_a_fill_onto_red_rows_alone_is_told_apart_from_one_changing_nothing()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")),
            Board(Component("U1", "CPU"), Component("U4", "PLA")));
        BoardTableSheet sheet = Components(document);
        BoardTableCell top = Row(sheet, "U1").Cells[NameColumn];
        BoardTableCell red2 = Row(sheet, "U2", deleted: true).Cells[NameColumn];
        BoardTableCell red3 = Row(sheet, "U3", deleted: true).Cells[NameColumn];
        BoardTableCell added = Row(sheet, "U4").Cells[NameColumn];

        Assert.True(sheet.OnlyDeletedRows([red2, red3]));
        Assert.False(sheet.OnlyDeletedRows([top, red2]));
        Assert.False(sheet.OnlyDeletedRows([top]));

        Assert.True(sheet.FillsOnlyDeletedRows([top, red2, red3]));
        Assert.False(sheet.FillsOnlyDeletedRows([top, red2, added]));
        Assert.False(sheet.FillsOnlyDeletedRows([top]));

        Assert.Equal(0, sheet.FillDown([top, red2, red3]));
        Assert.False(document.History.CanUndo);
    }

    [Fact]
    public void One_cell_alone_has_nothing_to_fill_and_leaves_no_undo_step()
    {
        BoardTableDocument document = Draft(Component("U1", "CPU"));
        BoardTableSheet sheet = Components(document);

        Assert.Equal(0, sheet.FillDown([Row(sheet, "U1").Cells[NameColumn]]));
        Assert.False(document.History.CanUndo);

        // And the history is not left grouping: the next edit is a step of its own.
        Row(sheet, "U1").Cells[NameColumn].Text = "CPU 6510";
        Row(sheet, "U1").Cells[LabelColumn].Text = "U7";
        Assert.Equal(2, document.History.UndoCount);
    }
}
