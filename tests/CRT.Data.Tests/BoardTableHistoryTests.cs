using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Undo and redo in the Drafts tab's table editor - Ctrl+Z / Ctrl+Y (owner request,
// 2026-09-24).
//
// Every step is a snapshot of one sheet's live rows (see BoardTableHistory's header), so these
// tests go through each KIND of change - a cell, an insert, a delete, a restore, a move, a revert
// - and check that undo lands on exactly the table as it was: the same rows, the same values, the
// same colours and ghosts, and the same "unsaved" flag. Redo must land on exactly the table as it
// was after the change.
// ###########################################################################################
public sealed class BoardTableHistoryTests
{
    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, TechnicalNameOrValue = "x" };

    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    private static BoardTableSheet Components(BoardTableDocument document) =>
        document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

    private static int Column(string column) =>
        BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(column);

    private static string Label(BoardTableRow row) => row.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text;

    private static BoardTableCell FriendlyName(BoardTableRow row) => row.Cells[Column(BoardWorkbookSchema.ColFriendlyName)];

    private static BoardTableRow Row(BoardTableSheet sheet, string label, bool deleted = false) =>
        sheet.Rows.Single(row => row.IsDeleted == deleted && Label(row) == label);

    // What the sheet looks like: each row's label, deleted-ness and state, in order.
    private static List<string> Picture(BoardTableSheet sheet) =>
        sheet.Rows.Select(row => $"{(row.IsDeleted ? "-" : "")}{Label(row)}:{row.State}").ToList();

    private static BoardTableDocument Open() =>
        BoardTableDocument.Create(
            published: Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")),
            draft: Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")));

    // ------------------------------------------------------------------ Each kind of change

    [Fact]
    public void Undo_puts_a_cell_back_with_its_colour_and_redo_does_the_edit_again()
    {
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);
        BoardTableRow u1 = Row(sheet, "U1");

        FriendlyName(u1).Text = "CPU 6510";
        sheet.Refresh();
        Assert.Equal(BoardTableCellState.Modified, FriendlyName(u1).State);

        BoardTableHistoryResult? undone = document.History.Undo();

        Assert.Equal("CPU", FriendlyName(u1).Text);
        Assert.Equal(BoardTableCellState.Unchanged, FriendlyName(u1).State);
        Assert.Equal(0, sheet.ChangeCount);

        // It says where: this sheet, this row, this column - so the cursor can go there.
        Assert.NotNull(undone);
        Assert.Same(sheet, undone.Sheet);
        Assert.Same(u1, undone.Row);
        Assert.Equal(Column(BoardWorkbookSchema.ColFriendlyName), undone.Column);

        document.History.Redo();

        Assert.Equal("CPU 6510", FriendlyName(u1).Text);
        Assert.Equal(BoardTableCellState.Modified, FriendlyName(u1).State);
        Assert.Equal(1, sheet.ChangeCount);
    }

    [Fact]
    public void Undo_removes_an_inserted_row_and_redo_brings_back_the_SAME_row()
    {
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);
        BoardTableRow u1 = Row(sheet, "U1");

        BoardTableRow inserted = sheet.InsertRow(u1);
        inserted.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text = "U9";
        inserted.Cells[Column(BoardWorkbookSchema.ColTechnicalNameOrValue)].Text = "x";
        sheet.Refresh();

        document.History.Undo();
        document.History.Undo();
        BoardTableHistoryResult? undone = document.History.Undo();

        Assert.DoesNotContain(inserted, sheet.Rows);
        Assert.Equal(["U1:Unchanged", "U2:Unchanged", "U3:Unchanged"], Picture(sheet));

        // The cursor goes back to the row it was on when "Insert row" was pressed.
        Assert.Same(u1, undone!.Row);
        Assert.Equal(-1, undone.Column);

        document.History.Redo();
        document.History.Redo();
        BoardTableHistoryResult? redone = document.History.Redo();

        Assert.Same(inserted, sheet.Rows[1]);
        Assert.Equal("U9", Label(inserted));
        Assert.Equal(BoardTableRowState.Added, inserted.State);
        Assert.Same(inserted, redone!.Row);
    }

    [Fact]
    public void Undo_brings_back_a_deleted_published_row_in_place_and_its_red_ghost_goes()
    {
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);
        BoardTableRow u2 = Row(sheet, "U2");

        sheet.DeleteRow(u2);
        Assert.Equal(["U1:Unchanged", "-U2:Deleted", "U3:Unchanged"], Picture(sheet));

        BoardTableHistoryResult? undone = document.History.Undo();

        Assert.Equal(["U1:Unchanged", "U2:Unchanged", "U3:Unchanged"], Picture(sheet));
        Assert.Same(u2, sheet.Rows[1]);
        Assert.Same(u2, undone!.Row);
        Assert.Equal(0, sheet.DeletedCount);

        document.History.Redo();

        Assert.Equal(["U1:Unchanged", "-U2:Deleted", "U3:Unchanged"], Picture(sheet));
    }

    [Fact]
    public void Undo_takes_back_a_restore_and_the_row_is_red_again()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            published: Board(Component("U1"), Component("U2")),
            draft: Board(Component("U1")));
        BoardTableSheet sheet = Components(document);

        BoardTableRow restored = sheet.RestoreRow(Row(sheet, "U2", deleted: true))!;
        Assert.Equal(["U1:Unchanged", "U2:Unchanged"], Picture(sheet));

        BoardTableHistoryResult? undone = document.History.Undo();

        Assert.Equal(["U1:Unchanged", "-U2:Deleted"], Picture(sheet));
        Assert.True(undone!.Row!.IsDeleted);

        document.History.Redo();

        Assert.Equal(["U1:Unchanged", "U2:Unchanged"], Picture(sheet));
        Assert.Same(restored, sheet.Rows[1]);
    }

    [Fact]
    public void Undo_takes_back_a_move_and_the_placement_it_switched_off()
    {
        // A row inserted this sitting is placed into its category on save UNLESS it was moved by
        // hand. Undoing the move must hand that decision back to the automatic placement.
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);

        BoardTableRow inserted = sheet.InsertRow(Row(sheet, "U3"));
        inserted.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text = "U0";
        inserted.Cells[Column(BoardWorkbookSchema.ColTechnicalNameOrValue)].Text = "x";
        sheet.Refresh();
        Assert.True(sheet.HasRowsToPlaceOnSave);

        sheet.MoveRow(inserted, 0);
        Assert.Same(inserted, sheet.Rows[0]);
        Assert.False(sheet.HasRowsToPlaceOnSave);

        BoardTableHistoryResult? undone = document.History.Undo();

        Assert.Same(inserted, sheet.Rows[3]);
        Assert.True(sheet.HasRowsToPlaceOnSave);
        Assert.Same(inserted, undone!.Row);

        document.History.Redo();

        Assert.Same(inserted, sheet.Rows[0]);
        Assert.False(sheet.HasRowsToPlaceOnSave);
    }

    [Fact]
    public void Undo_takes_back_a_revert_cell()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            published: Board(Component("U1", "CPU")),
            draft: Board(Component("U1", "CPU 6510")));
        BoardTableSheet sheet = Components(document);
        BoardTableCell cell = FriendlyName(Row(sheet, "U1"));

        Assert.True(sheet.RevertCell(cell));
        Assert.Equal("CPU", cell.Text);

        document.History.Undo();

        Assert.Equal("CPU 6510", cell.Text);
        Assert.Equal(BoardTableCellState.Modified, cell.State);
    }

    // ------------------------------------------------------------------ The stacks

    [Fact]
    public void Several_steps_undo_newest_first_across_sheets_and_say_which_sheet()
    {
        BoardTableDocument document = Open();
        BoardTableSheet components = Components(document);
        BoardTableSheet links = document.FindSheet(BoardWorkbookSchema.SheetBoardLinks)!;

        FriendlyName(Row(components, "U1")).Text = "one";
        BoardTableRow link = links.InsertRow(null);
        FriendlyName(Row(components, "U2")).Text = "two";

        Assert.Equal(3, document.History.UndoCount);

        Assert.Same(components, document.History.Undo()!.Sheet);
        Assert.Equal("VIC", FriendlyName(Row(components, "U2")).Text);

        Assert.Same(links, document.History.Undo()!.Sheet);
        Assert.DoesNotContain(link, links.Rows);

        Assert.Same(components, document.History.Undo()!.Sheet);
        Assert.Equal("CPU", FriendlyName(Row(components, "U1")).Text);

        Assert.Null(document.History.Undo());
        Assert.Equal(3, document.History.RedoCount);
    }

    [Fact]
    public void A_new_change_after_an_undo_throws_the_redo_away()
    {
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);

        FriendlyName(Row(sheet, "U1")).Text = "first";
        document.History.Undo();
        Assert.True(document.History.CanRedo);

        FriendlyName(Row(sheet, "U2")).Text = "second";

        Assert.False(document.History.CanRedo);
        Assert.Null(document.History.Redo());
        Assert.Equal("CPU", FriendlyName(Row(sheet, "U1")).Text);
    }

    [Fact]
    public void Nothing_to_undo_or_redo_on_a_table_just_opened()
    {
        BoardTableDocument document = Open();

        Assert.False(document.History.CanUndo);
        Assert.False(document.History.CanRedo);
        Assert.Null(document.History.Undo());
        Assert.Null(document.History.Redo());
    }

    [Fact]
    public void A_change_that_changes_nothing_records_no_step()
    {
        // Retyping a cell's own value, or a move onto the same place, is not something to undo -
        // an undo that visibly did nothing would look broken.
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);
        BoardTableRow u1 = Row(sheet, "U1");

        FriendlyName(u1).Text = "CPU";
        FriendlyName(u1).Text = " CPU ";
        sheet.MoveRow(u1, 0);
        sheet.RevertCell(FriendlyName(u1));

        Assert.Equal(0, document.History.UndoCount);
    }

    // ------------------------------------------------------------------ Unsaved changes

    [Fact]
    public void Undoing_every_change_leaves_nothing_unsaved_and_redo_makes_it_unsaved_again()
    {
        // "Save changes" should grey out again once everything has been taken back.
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);

        FriendlyName(Row(sheet, "U1")).Text = "one";
        sheet.DeleteRow(Row(sheet, "U3"));
        Assert.True(document.HasUnsavedChanges);

        document.History.Undo();
        Assert.True(document.HasUnsavedChanges);

        document.History.Undo();
        Assert.False(document.HasUnsavedChanges);

        document.History.Redo();
        Assert.True(document.HasUnsavedChanges);
    }

    [Fact]
    public void After_a_save_the_saved_state_is_the_clean_one()
    {
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);

        FriendlyName(Row(sheet, "U1")).Text = "one";
        document.MarkSaved();
        Assert.False(document.HasUnsavedChanges);

        // Back past the save: the table now differs from what was saved.
        document.History.Undo();
        Assert.True(document.HasUnsavedChanges);

        // And forward onto it again: clean.
        document.History.Redo();
        Assert.False(document.HasUnsavedChanges);
    }

    [Fact]
    public void Once_the_saved_state_is_thrown_away_no_undo_can_claim_to_reach_it()
    {
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);

        FriendlyName(Row(sheet, "U1")).Text = "one";
        document.MarkSaved();
        document.History.Undo();

        // A different change from here: the saved "one" is on the redo stack, and goes.
        FriendlyName(Row(sheet, "U2")).Text = "two";
        document.History.Undo();

        // Back to the opening values - but the file holds "one", so this is NOT clean.
        Assert.Equal("CPU", FriendlyName(Row(sheet, "U1")).Text);
        Assert.True(document.HasUnsavedChanges);
    }

    [Fact]
    public void The_oldest_steps_are_dropped_past_the_limit_and_the_start_is_then_not_clean()
    {
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);
        BoardTableCell cell = FriendlyName(Row(sheet, "U1"));

        for (int i = 0; i < BoardTableHistory.MaxSteps + 5; i++)
        {
            cell.Text = $"value {i}";
        }

        Assert.Equal(BoardTableHistory.MaxSteps, document.History.UndoCount);

        while (document.History.Undo() is not null)
        {
        }

        // As far back as the history reaches - five edits short of the original "CPU".
        Assert.Equal("value 4", cell.Text);
        Assert.True(document.HasUnsavedChanges);
    }

    [Fact]
    public void Undo_keeps_the_counts_agreeing_with_a_fresh_table_of_the_same_rows()
    {
        // The strongest check that nothing is left half-restored: after a mixed run of changes and
        // undos, the sheet must look exactly as a table built from scratch on its rows would.
        BoardTableDocument document = Open();
        BoardTableSheet sheet = Components(document);

        FriendlyName(Row(sheet, "U1")).Text = "changed";
        sheet.DeleteRow(Row(sheet, "U2"));
        BoardTableRow added = sheet.InsertRow(Row(sheet, "U3"));
        added.Cells[Column(BoardWorkbookSchema.ColBoardLabel)].Text = "U7";
        added.Cells[Column(BoardWorkbookSchema.ColTechnicalNameOrValue)].Text = "x";
        sheet.Refresh();

        // Back past the typing and the insert, leaving the edit and the delete.
        document.History.Undo();
        document.History.Undo();
        document.History.Undo();
        Assert.Equal(["U1:Modified", "-U2:Deleted", "U3:Unchanged"], Picture(sheet));

        BoardTableDocument fresh = BoardTableDocument.Create(
            published: Board(Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID")),
            draft: document.ApplyTo(new BoardData()));

        Assert.Equal(Picture(Components(fresh)), Picture(sheet));
        Assert.Equal(Components(fresh).ChangeCount, sheet.ChangeCount);
    }
}
