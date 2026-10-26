using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // COPY AND PASTE OF BLOCKS, FILL DOWN, AND FIND & REPLACE (CRT 4.0, 2026-10-25 - the plan's
    // step 0: "the table view gains block copy and paste, fill down, and find & replace"; with the
    // board data out of Excel, this table is where every edit is made, and without these dropping
    // Excel would be a step backwards).
    //
    //   - Ctrl+C copies the selected cells, as a spreadsheet reads them (BoardTableClipboard.ForCells);
    //   - Ctrl+V pastes a block from the selection's top-left cell down the rows on screen, adding
    //     rows at the end when it runs out (BoardTableSheet.PasteBlock), or ONE value onto every
    //     selected cell (PasteIntoCells);
    //   - Ctrl+D or "Fill down" copies each column's top selected cell into the ones below it, or -
    //     on one row - the row above into it (BoardTableSheet.FillDown);
    //   - Ctrl+H or "Replace" opens the replace bar: every occurrence of a text, ignoring capitals,
    //     in the rows the table shows, on this sheet or on every sheet (BoardTableDocument.ReplaceAll).
    //   - Ctrl+F puts the cursor in the search box.
    //
    // Each is ONE undo step, and says what it did on the status line (BoardTableEditWording). The
    // rules are all in CRT.Data; this part only reads the selection and hands it over. A read-only
    // table offers none of the changes - copying stays.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // ###########################################################################################
        // The cells selected, each once, in the sheet's order and then the columns' - only data cells
        // of rows on screen. Empty with none.
        // ###########################################################################################
        internal List<BoardTableCell> SelectedTableCells()
        {
            if (this.thisCurrentSheet is not { } sheet)
            {
                return [];
            }

            var position = new Dictionary<BoardTableRow, int>(ReferenceEqualityComparer.Instance);
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                position[sheet.Rows[i]] = i;
            }

            var cells = new List<BoardTableCell>();

            foreach (DataGridCellInfo info in this.TableGrid.SelectedCells)
            {
                if (info.Item is BoardTableRow row &&
                    position.ContainsKey(row) &&
                    (this.thisView is null || this.thisView.IndexOf(row) >= 0) &&
                    info.Column?.Tag is int column &&
                    column >= 0 &&
                    column < row.Cells.Count)
                {
                    cells.Add(row.Cells[column]);
                }
            }

            return cells
                .Distinct(ReferenceEqualityComparer.Instance)
                .Cast<BoardTableCell>()
                .OrderBy(cell => position[cell.Row])
                .ThenBy(cell => cell.ColumnIndex)
                .ToList();
        }

        // Ctrl+C: the selected cells - or, with one or none selected, the cell with the dashed frame.
        internal string? CopySelectionText()
        {
            List<BoardTableCell> cells = this.SelectedTableCells();

            if (cells.Count > 1)
            {
                return BoardTableClipboard.ForCells(cells, this.ShownRows());
            }

            BoardTableCell? cell = cells.Count == 1 ? cells[0] : this.CurrentCell;

            return cell is null ? null : BoardTableClipboard.ForSingleCell(cell.Text);
        }

        // ###########################################################################################
        // Ctrl+V. One value onto several selected cells fills them all; anything else lands from the
        // selection's top-left cell - the framed cell when only it is selected. Nothing on a red row,
        // in a read-only table, or from a clipboard with no text - each of the last two said on the
        // status line. True when something was pasted.
        // ###########################################################################################
        internal bool Paste(string? clipboardText)
        {
            if (this.thisIsReadOnly || this.thisCurrentSheet is not { } sheet || this.CurrentCell is not { } current)
            {
                return false;
            }

            this.TableGrid.CommitEdit();

            IReadOnlyList<IReadOnlyList<string>> block = BoardTableClipboard.ReadBlock(clipboardText);
            if (block.Count == 0)
            {
                this.ShowStatus(BoardTableEditWording.NothingToPaste);
                return false;
            }

            List<BoardTableCell> selected = this.SelectedTableCells();

            if (block.Count == 1 && block[0].Count == 1 && selected.Count > 1)
            {
                if (sheet.OnlyDeletedRows(selected))
                {
                    this.ShowStatus(BoardTableEditWording.PasteOnDeletedRows);
                    return false;
                }

                this.ShowStatus(BoardTableEditWording.PastedIntoCells(sheet.PasteIntoCells(selected, block[0][0])));
                this.UpdateToolbar();
                return true;
            }

            BoardTableCell anchor = selected.Count > 1
                ? selected[0].Row.Cells[selected.Min(cell => cell.ColumnIndex)]
                : current;

            if (anchor.Row.IsDeleted)
            {
                this.ShowStatus(BoardTableEditWording.PasteOnDeletedRow);
                return false;
            }

            BoardTablePasteResult result = sheet.PasteBlock(anchor, block, this.ShownRows(), mayAddRows: !this.IsNarrowed);

            this.ShowStatus(BoardTableEditWording.Pasted(result));

            if (result.Rows.Count > 1 || result.ColumnCount > 1)
            {
                this.SelectPasted(result, anchor);
            }

            this.UpdateToolbar();

            return result.Pasted;
        }

        // The cells a paste wrote, selected, with the cursor back on the anchor - so what went in is
        // in view, and a Ctrl+D or a Ctrl+C straight after acts on it.
        private void SelectPasted(BoardTablePasteResult result, BoardTableCell anchor)
        {
            this.SelectCell(anchor.Row, anchor.ColumnIndex);

            var cells = new List<DataGridCellInfo>();

            foreach (BoardTableRow row in result.Rows)
            {
                int rowIndex = this.thisView?.IndexOf(row) ?? -1;

                for (int column = result.FirstColumn; rowIndex >= 0 && column < result.FirstColumn + result.ColumnCount; column++)
                {
                    if (this.TableGrid.Columns.FirstOrDefault(c => c.Tag is int index && index == column) is { } gridColumn)
                    {
                        cells.Add(new DataGridCellInfo(row, gridColumn, rowIndex, gridColumn.DisplayIndex, true));
                    }
                }
            }

            if (cells.Count > 1)
            {
                this.TableGrid.SelectedCells = cells;
            }
        }

        // ###########################################################################################
        // Ctrl+D / "Fill down". Several rows selected: each column's top cell fills the ones below.
        // One row: the row shown above fills it, as a spreadsheet's Ctrl+D does.
        // ###########################################################################################
        internal void FillDown()
        {
            if (this.thisIsReadOnly || this.thisCurrentSheet is not { } sheet)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            List<BoardTableCell> cells = this.SelectedTableCells();

            if (cells.Count == 0 && this.CurrentCell is { } current)
            {
                cells = [current];
            }

            if (cells.Count == 0)
            {
                return;
            }

            if (cells.All(cell => ReferenceEquals(cell.Row, cells[0].Row)))
            {
                if (this.RowShownAbove(cells[0].Row) is not { } above)
                {
                    this.ShowStatus(BoardTableEditWording.NothingAbove);
                    return;
                }

                cells = [.. cells.Select(cell => above.Cells[cell.ColumnIndex]), .. cells];
            }

            this.ShowStatus(sheet.FillsOnlyDeletedRows(cells)
                ? BoardTableEditWording.FillOnDeletedRows
                : BoardTableEditWording.Filled(sheet.FillDown(cells)));
            this.UpdateToolbar();
        }

        // The row shown straight above `row`, or null at the top.
        private BoardTableRow? RowShownAbove(BoardTableRow row)
        {
            if (this.thisView is null)
            {
                int index = this.thisCurrentSheet?.Rows.IndexOf(row) ?? -1;
                return index > 0 ? this.thisCurrentSheet!.Rows[index - 1] : null;
            }

            int at = this.thisView.IndexOf(row);
            return at > 0 ? this.thisView.GetItemAt(at - 1) as BoardTableRow : null;
        }

        // Whether "Fill down" has anything to do: a live row below the top of a selection, or one row
        // with a row shown above it.
        private bool CanFillDown()
        {
            if (this.thisIsReadOnly || this.thisCurrentSheet is null)
            {
                return false;
            }

            List<BoardTableRow> rows = this.SelectedRows();

            return rows.Count > 1
                ? rows.Skip(1).Any(row => !row.IsDeleted)
                : rows.Count == 1 && !rows[0].IsDeleted && this.RowShownAbove(rows[0]) is not null;
        }

        // ###########################################################################################
        // The replace bar. Opened by "Replace" or Ctrl+H, closed by its own "Close" or Escape; Enter in
        // the "Find" box moves on to "Replace with", Enter there replaces.
        // ###########################################################################################
        internal void OpenReplaceBar()
        {
            if (this.thisIsReadOnly || this.thisDocument is null)
            {
                return;
            }

            this.TableGrid.CommitEdit();
            this.ReplaceBar.IsVisible = true;
            this.ReplaceFindBox.Focus();
            this.ReplaceFindBox.SelectAll();
        }

        internal bool IsReplaceBarOpen => this.ReplaceBar.IsVisible;

        private void CloseReplaceBar()
        {
            this.ReplaceBar.IsVisible = false;
            this.FocusGrid();
        }

        // "Replace all": the rows the table shows - the pills and the search as they are - on this
        // sheet, or on every sheet with "On every sheet" ticked.
        internal void ReplaceAll()
        {
            if (this.thisIsReadOnly || this.thisDocument is not { } document || this.thisCurrentSheet is not { } sheet)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            string find = this.ReplaceFindBox.Text ?? string.Empty;

            if (find.Length == 0)
            {
                this.ShowStatus(BoardTableEditWording.NothingToFind);
                this.ReplaceFindBox.Focus();
                return;
            }

            bool everySheet = this.ReplaceEverySheetBox.IsChecked == true;
            BoardTableRowKinds kinds = this.thisFilter;
            BoardTableSearch search = this.thisSearch;

            BoardTableReplaceResult result = document.ReplaceAll(
                find,
                this.ReplaceWithBox.Text,
                everySheet ? document.Sheets : [sheet],
                row => BoardTableRowFilter.Shows(kinds, row) && search.Shows(row));

            this.ShowStatus(BoardTableEditWording.Replaced(result, find, everySheet));
            this.UpdateSheetTabs();
            this.UpdateToolbar();
        }

        private void WireBlockEdits()
        {
            this.ReplaceFindBox.KeyDown += this.OnReplaceBoxKeyDown;
            this.ReplaceWithBox.KeyDown += this.OnReplaceBoxKeyDown;
        }

        private void OnReplaceBoxKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                this.CloseReplaceBar();
            }
            else if (e.Key == Key.Enter && ReferenceEquals(sender, this.ReplaceFindBox))
            {
                e.Handled = true;
                this.ReplaceWithBox.Focus();
            }
            else if (e.Key == Key.Enter)
            {
                e.Handled = true;
                this.ReplaceAll();
            }
        }

        private void OnFillDownClick(object? sender, RoutedEventArgs e) => this.ThenFocusGrid(this.FillDown);

        private void OnReplaceClick(object? sender, RoutedEventArgs e)
        {
            if (this.ReplaceBar.IsVisible)
            {
                this.CloseReplaceBar();
            }
            else
            {
                this.OpenReplaceBar();
            }
        }

        private void OnReplaceAllClick(object? sender, RoutedEventArgs e) => this.ReplaceAll();

        private void OnCloseReplaceClick(object? sender, RoutedEventArgs e) => this.CloseReplaceBar();

        // The buttons of this part, as UpdateToolbar sets the rest: none of them in a read-only table.
        private void UpdateBlockEditButtons()
        {
            this.FillDownButton.IsVisible = !this.thisIsReadOnly;
            this.FillDownButton.IsEnabled = this.CanFillDown();
            this.ReplaceButton.IsVisible = !this.thisIsReadOnly;
            this.ReplaceButton.IsEnabled = this.thisCurrentSheet is not null;

            if (this.thisIsReadOnly)
            {
                this.ReplaceBar.IsVisible = false;
            }
        }

        // ###########################################################################################
        // A key with the platform's command modifier alone - Ctrl, or Cmd on macOS - for the shortcuts
        // the platform's hotkey list has no entry for (fill down, find, replace). Ctrl where the
        // platform does not say, as on the headless one the tests run on.
        // ###########################################################################################
        private static bool IsCommand(KeyEventArgs e, Key key)
        {
            KeyModifiers command = Application.Current?.PlatformSettings?.HotkeyConfiguration.CommandModifiers ?? KeyModifiers.Control;

            return e.Key == key && e.KeyModifiers == command;
        }
    }
}
