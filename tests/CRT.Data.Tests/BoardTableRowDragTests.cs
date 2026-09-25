using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Dragging a row in the Drafts tab's table editor with a live placeholder (maintainer request,
// 2026-09-24: "like moving an image in the worklog"). The row moves while it is dragged, into the
// place of the row under the pointer; these pin where it lands, that a red ghost is never a place
// to land, and that a whole drag is one undo step - or none, when it ends where it began.
// ###########################################################################################
public sealed class BoardTableRowDragTests
{
    private static ComponentEntry Component(string label) =>
        new() { BoardLabel = label, TechnicalNameOrValue = "x" };

    private static BoardData Board(params string[] labels) =>
        new() { Components = [.. labels.Select(Component)] };

    private static BoardTableSheet Components(BoardTableDocument document) =>
        document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

    private static int LabelColumn => BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel);

    private static BoardTableRow Row(BoardTableSheet sheet, string label, bool deleted = false) =>
        sheet.Rows.Single(row => row.IsDeleted == deleted && row.Cells[LabelColumn].Text == label);

    // Labels in order, a deleted ghost prefixed with "-".
    private static List<string> Order(BoardTableSheet sheet) =>
        sheet.Rows.Select(row => (row.IsDeleted ? "-" : "") + row.Cells[LabelColumn].Text).ToList();

    [Fact]
    public void Dragging_down_onto_a_row_takes_its_place_and_it_moves_up_one()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2", "U3", "U4"));
        BoardTableSheet sheet = Components(document);

        BoardTableRowDrag drag = sheet.BeginRowDrag(Row(sheet, "U1"))!;

        Assert.True(drag.MoveOnto(Row(sheet, "U2")));
        Assert.Equal(["U2", "U1", "U3", "U4"], Order(sheet));

        Assert.True(drag.MoveOnto(Row(sheet, "U4")));
        Assert.Equal(["U2", "U3", "U4", "U1"], Order(sheet));

        Assert.True(drag.Finish());
    }

    [Fact]
    public void Dragging_up_onto_a_row_takes_its_place_and_it_moves_down_one()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2", "U3"));
        BoardTableSheet sheet = Components(document);

        BoardTableRowDrag drag = sheet.BeginRowDrag(Row(sheet, "U3"))!;
        Assert.True(drag.MoveOnto(Row(sheet, "U1")));

        Assert.Equal(["U3", "U1", "U2"], Order(sheet));
        drag.Finish();
    }

    [Fact]
    public void Hovering_the_dragged_row_itself_changes_nothing()
    {
        // After each move the dragged row IS under the pointer - that must be a resting state,
        // not another move.
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2"));
        BoardTableSheet sheet = Components(document);

        BoardTableRowDrag drag = sheet.BeginRowDrag(Row(sheet, "U1"))!;

        Assert.False(drag.MoveOnto(Row(sheet, "U1")));
        drag.Finish();
    }

    [Fact]
    public void A_red_deleted_row_is_not_a_place_to_drop_onto_and_cannot_be_dragged()
    {
        // Where a ghost shows is worked out from the published order after every move, so moving
        // "onto" one would be put straight back by the refresh - and flicker for as long as the
        // pointer stayed there.
        BoardTableDocument document = BoardTableDocument.Create(
            published: Board("U1", "U2", "U3"),
            draft: Board("U1", "U3"));
        BoardTableSheet sheet = Components(document);
        Assert.Equal(["U1", "-U2", "U3"], Order(sheet));

        Assert.Null(sheet.BeginRowDrag(Row(sheet, "U2", deleted: true)));

        BoardTableRowDrag drag = sheet.BeginRowDrag(Row(sheet, "U3"))!;
        Assert.False(drag.MoveOnto(Row(sheet, "U2", deleted: true)));
        Assert.Equal(["U1", "-U2", "U3"], Order(sheet));

        // Past it, onto the live row above, is fine.
        Assert.True(drag.MoveOnto(Row(sheet, "U1")));
        Assert.Equal("U3", Order(sheet)[0]);
        drag.Finish();
    }

    [Fact]
    public void Stepping_past_the_edge_moves_one_live_row_at_a_time()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2", "U3"));
        BoardTableSheet sheet = Components(document);

        BoardTableRowDrag drag = sheet.BeginRowDrag(Row(sheet, "U1"))!;

        Assert.True(drag.Step(up: false));
        Assert.True(drag.Step(up: false));
        Assert.False(drag.Step(up: false));

        Assert.Equal(["U2", "U3", "U1"], Order(sheet));
        drag.Finish();
    }

    [Fact]
    public void A_whole_drag_is_ONE_undo_step_however_many_rows_it_passed()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2", "U3", "U4"));
        BoardTableSheet sheet = Components(document);

        BoardTableRowDrag drag = sheet.BeginRowDrag(Row(sheet, "U1"))!;
        drag.MoveOnto(Row(sheet, "U2"));
        drag.MoveOnto(Row(sheet, "U3"));
        drag.MoveOnto(Row(sheet, "U4"));
        drag.Finish();

        Assert.Equal(1, document.History.UndoCount);

        document.History.Undo();
        Assert.Equal(["U1", "U2", "U3", "U4"], Order(sheet));
    }

    [Fact]
    public void Changes_after_a_drag_are_their_own_steps_again()
    {
        // An unfinished group would swallow every later change into the drag's step.
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2"));
        BoardTableSheet sheet = Components(document);

        BoardTableRowDrag drag = sheet.BeginRowDrag(Row(sheet, "U1"))!;
        drag.MoveOnto(Row(sheet, "U2"));
        drag.Finish();

        Row(sheet, "U1").Cells[LabelColumn + 1].Text = "one";
        Row(sheet, "U2").Cells[LabelColumn + 1].Text = "two";

        Assert.Equal(3, document.History.UndoCount);
    }

    [Fact]
    public void A_drag_that_ends_where_it_began_leaves_no_step_nothing_unsaved_and_its_placement_intact()
    {
        // An inserted row is placed into its category on save UNLESS moved by hand. Taking it for a
        // trip and back is not a move, and must not cost it that.
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2", "U3"));
        BoardTableSheet sheet = Components(document);

        BoardTableRow inserted = sheet.InsertRow(Row(sheet, "U3"));
        inserted.Cells[LabelColumn].Text = "U0";
        sheet.Refresh();
        document.MarkSaved();
        int steps = document.History.UndoCount;
        Assert.True(sheet.HasRowsToPlaceOnSave);

        BoardTableRowDrag drag = sheet.BeginRowDrag(inserted)!;
        drag.MoveOnto(Row(sheet, "U1"));
        Assert.True(document.HasUnsavedChanges);
        drag.MoveOnto(Row(sheet, "U3"));

        Assert.False(drag.Finish());

        Assert.Equal(["U1", "U2", "U3", "U0"], Order(sheet));
        Assert.Equal(steps, document.History.UndoCount);
        Assert.False(document.HasUnsavedChanges);
        Assert.True(sheet.HasRowsToPlaceOnSave);
    }

    [Fact]
    public void A_finished_drag_moves_nothing_more()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2", "U3"));
        BoardTableSheet sheet = Components(document);

        BoardTableRowDrag drag = sheet.BeginRowDrag(Row(sheet, "U1"))!;
        drag.Finish();

        Assert.False(drag.MoveOnto(Row(sheet, "U3")));
        Assert.False(drag.Step(up: false));
        Assert.False(drag.Finish());
        Assert.Equal(["U1", "U2", "U3"], Order(sheet));
    }

    // ------------------------------------------------------------------ Insert row above

    [Fact]
    public void Insert_row_above_puts_the_empty_row_straight_before_the_selected_one()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2"));
        BoardTableSheet sheet = Components(document);

        BoardTableRow inserted = sheet.InsertRowAbove(Row(sheet, "U2"));

        Assert.Same(inserted, sheet.Rows[1]);
        Assert.True(inserted.IsBlank);

        // With nothing selected, at the very top - the "above" twin of "below" going to the end.
        BoardTableRow top = sheet.InsertRowAbove(null);
        Assert.Same(top, sheet.Rows[0]);
    }

    [Fact]
    public void Undoing_insert_row_above_takes_the_cursor_back_to_the_row_it_was_on()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board("U1", "U2"));
        BoardTableSheet sheet = Components(document);
        BoardTableRow u2 = Row(sheet, "U2");

        sheet.InsertRowAbove(u2);
        BoardTableHistoryResult? undone = document.History.Undo();

        Assert.Equal(["U1", "U2"], Order(sheet));
        Assert.Same(u2, undone!.Row);
    }
}
