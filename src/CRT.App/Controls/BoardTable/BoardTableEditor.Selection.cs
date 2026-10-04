using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // SEVERAL ROWS SELECTED, AND DELETED IN ONE GO (owner request, 2026-10-02: "how difficult would
    // it be to mark multiple rows and then delete those in one go? E.g. ... I would like to mark all
    // 8 pinout rows and delete those").
    //
    // The grid selects CELLS (SelectionMode Extended): Shift+click takes every cell from the first
    // click to it, Ctrl+click adds one or takes it out, and a row is selected when any of its cells
    // is. "Delete row" then deletes every selected row - BoardTableSheet.DeleteRows, one change and
    // one undo step - and says how many ("Delete 8 rows"). Red rows among them are skipped, already
    // deleted; with only red rows selected the button is greyed out. Copy, paste and dragging a row
    // by its grip still act on one cell or row.
    //
    // *** THE DELETE KEY DELETES NO ROW (agreed, 2026-10-02: "No"). *** In Excel it empties cells,
    // and a row delete on a reflex key press is easy to miss. The grid's CanUserDeleteRows is off
    // in the markup - it was on by default, and Delete on a selected cell took the row straight out
    // of the sheet's list behind the table's back: no red row, no undo step.
    //
    // *** A SELECTION THE CURSOR HAS LEFT IS NOT THE ONE ACTED ON. *** The cursor - the cell with
    // the dashed frame - is moved from code too: undo puts it on the row it brought back, a delete
    // on the place of the first row deleted, Tab on the next cell. SelectCell makes that one cell
    // the whole selection whenever it does, so "Delete row" never takes rows selected before.
    // Ctrl+click on a selected cell takes it out of the selection while the frame moves onto it;
    // the rows still selected are then the ones acted on, as their veil shows.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // The grid class while more than one cell is selected - each selected cell then carries a
        // light veil (BoardTable_Selection_Fill), which a single cell does not: its frame says it.
        internal const string SeveralSelectedClass = "SeveralSelected";

        // The rows selected, in the sheet's order - only rows on screen. With no cell selected, the
        // cursor's row; empty with no cursor either.
        internal List<BoardTableRow> SelectedRows()
        {
            if (this.thisCurrentSheet is null)
            {
                return [];
            }

            var selected = new HashSet<BoardTableRow>(
                this.TableGrid.SelectedCells.Select(cell => cell.Item).OfType<BoardTableRow>(),
                ReferenceEqualityComparer.Instance);

            List<BoardTableRow> rows = this.thisCurrentSheet.Rows
                .Where(row => selected.Contains(row) && (this.thisView is null || this.thisView.IndexOf(row) >= 0))
                .ToList();

            return rows.Count > 0 || this.CurrentRow is not { } current ? rows : [current];
        }

        private void WireSelection() =>
            this.TableGrid.SelectedCellsChanged += (_, _) => this.UpdateToolbar();

        // "Delete row": every selected row that is not already deleted.
        internal void DeleteRow()
        {
            if (this.thisCurrentSheet is null)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            List<BoardTableRow> rows = this.SelectedRows().Where(row => !row.IsDeleted).ToList();
            if (rows.Count == 0)
            {
                return;
            }

            BoardTableSheet sheet = this.thisCurrentSheet;
            int index = sheet.Rows.IndexOf(rows[0]);
            int column = this.CurrentCell?.ColumnIndex ?? 0;

            if (sheet.DeleteRows(rows, out BoardTableDeletedWith? deletedWith) > 0 && this.RowShownAtOrAfter(index) is { } next)
            {
                this.SelectCell(next, column);
            }

            // A deleted component takes its rows on other sheets and its highlights with it
            // (2026-09-25) - said here, since those sheets are not the one on screen.
            if (deletedWith is not null)
            {
                this.ShowStatus(deletedWith.Describe());
            }

            this.UpdateToolbar();
        }

        // Where the cursor goes after a delete: the row now at `index` - the first deleted row's red
        // row, or the row that moved up into its place - or, when a picked pill hides that one, the
        // next row on screen after it, else the last one before it.
        private BoardTableRow? RowShownAtOrAfter(int index)
        {
            if (this.thisCurrentSheet is not { } sheet || sheet.Rows.Count == 0)
            {
                return null;
            }

            index = Math.Clamp(index, 0, sheet.Rows.Count - 1);

            bool Shown(BoardTableRow row) => this.thisView is null || this.thisView.IndexOf(row) >= 0;

            return sheet.Rows.Skip(index).FirstOrDefault(Shown)
                ?? sheet.Rows.Take(index).LastOrDefault(Shown);
        }

        // What "Delete row" says, and whether it can: the rows it would delete.
        private void UpdateDeleteButton()
        {
            int rows = this.SelectedRows().Count(row => !row.IsDeleted);

            this.DeleteRowButton.IsEnabled = rows > 0;
            this.DeleteRowButton.Content = rows > 1 ? $"Delete {rows} rows" : "Delete row";
            this.TableGrid.Classes.Set(BoardTableEditor.SeveralSelectedClass, this.TableGrid.SelectedCells.Count > 1);
        }
    }
}
