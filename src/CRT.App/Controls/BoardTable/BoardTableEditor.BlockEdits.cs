using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // COPY AND PASTE OF BLOCKS, AND FIND & REPLACE (CRT 4.0, 2026-10-25 - the plan's step 0: "the
    // table view gains block copy and paste, fill down, and find & replace"; with the board data out
    // of Excel, this table is where every edit is made, and without these dropping Excel would be a
    // step backwards). Fill down went on 2026-11-07 (owner request: "I do not understand what "Fill
    // down" does ... please remove this functionality entirely").
    //
    //   - Ctrl+C copies the selected cells, as a spreadsheet reads them (BoardTableClipboard.ForCells);
    //   - Ctrl+V pastes a block from the selection's top-left cell down the rows on screen, adding
    //     rows at the end when it runs out (BoardTableSheet.PasteBlock), or ONE value onto every
    //     selected cell (PasteIntoCells);
    //   - Ctrl+H or "Replace" opens the replace bar: every occurrence of a text, ignoring capitals,
    //     in the rows the table shows on this sheet (BoardTableDocument.ReplaceAll) - its "Find"
    //     filtering and marking the rows as it is typed, the replacement drawn in its place;
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

            BoardTablePasteResult result = sheet.PasteBlock(anchor, block, this.ShownRows(), mayAddRows: !this.IsNarrowed, this.ColumnsShownNow());

            this.ShowStatus(BoardTableEditWording.Pasted(result));

            if (result.Rows.Count > 1 || result.ColumnCount > 1)
            {
                this.SelectPasted(result, anchor);
            }

            this.UpdateToolbar();

            return result.Pasted;
        }

        // The cells a paste wrote, selected, with the cursor back on the anchor - so what went in is
        // in view, and a Ctrl+C straight after acts on it.
        private void SelectPasted(BoardTablePasteResult result, BoardTableCell anchor)
        {
            this.SelectCell(anchor.Row, anchor.ColumnIndex);

            var cells = new List<DataGridCellInfo>();

            List<int> columns = this.ColumnsShownNow().SkipWhile(column => column != result.FirstColumn).Take(result.ColumnCount).ToList();

            foreach (BoardTableRow row in result.Rows)
            {
                int rowIndex = this.thisView?.IndexOf(row) ?? -1;

                foreach (int column in columns)
                {
                    if (rowIndex >= 0 && this.TableGrid.Columns.FirstOrDefault(c => c.Tag is int index && index == column) is { } gridColumn)
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
        // The replace bar. Opened by "Replace" or Ctrl+H, closed by its own "Cancel" or Escape; Enter in
        // the "Find" box moves on to "Replace with", Enter there replaces.
        //
        // *** WHILE IT IS OPEN, "FIND" IS THE TABLE'S FILTER (owner request, 2026-11-07: "I want that 'Find' works
        // in same way as the filtering, where it will filter and highlight my search string ... only search for the
        // string as-is"). *** The rows holding the text as typed are shown and it is marked (BoardTableSearch.Literal),
        // what "Replace with" holds drawn in its place in the replacement's colors - and the search box is emptied and
        // off ("do remove the possibility to also filter the table, as that could conflict"), back once the bar
        // closes, the table then showing every row again. "Replace all" counts the places it will replace.
        // ###########################################################################################
        internal void OpenReplaceBar()
        {
            if (this.thisIsReadOnly || this.thisDocument is null)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            // What the table was filtered by is what "Find" starts with, the cursor after it and nothing selected (owner
            // request, 2026-11-07: 'it should not "mark all text" in the "Find" input field. It should just place (with
            // focus) the cursor at the end of the text, which should match the filtering text').
            string searched = this.SearchBox.Text ?? string.Empty;

            if (searched.Trim().Length > 0)
            {
                this.ReplaceFindBox.Text = searched;
            }

            this.ReplaceBar.IsVisible = true;
            this.SearchBox.Text = string.Empty;
            this.SearchBox.IsEnabled = false;
            this.ApplyReplaceFind();
            this.ReplaceFindBox.Focus();
            this.ReplaceFindBox.ClearSelection();
            this.ReplaceFindBox.CaretIndex = this.ReplaceFindBox.Text?.Length ?? 0;
        }

        internal bool IsReplaceBarOpen => this.ReplaceBar.IsVisible;

        // ###########################################################################################
        // The bar gone without a word - the table closed, or made read-only under it: the search box on again, and "Find"
        // no longer the table's filter (code review, 2026-11-07: made read-only - a submission's detail arriving after
        // Ctrl+H - the table went on showing only the rows "Find" held, the search box empty and nothing saying why).
        // ###########################################################################################
        private void LetGoOfReplaceBar()
        {
            if (!this.ReplaceBar.IsVisible)
            {
                return;
            }

            this.thisReplaceFindTimer?.Stop();
            this.ReplaceBar.IsVisible = false;
            this.SearchBox.IsEnabled = true;

            if (this.thisSearchIsLiteral)
            {
                this.ApplySearch(string.Empty);
            }
        }

        private void CloseReplaceBar()
        {
            this.thisReplaceFindTimer?.Stop();
            this.ReplaceBar.IsVisible = false;
            this.SearchBox.IsEnabled = true;

            // Every row again: the search box was empty all the while.
            this.ApplySearch(string.Empty);
            this.FocusGrid();
        }

        // ###########################################################################################
        // What "Find" holds, applied as the table's filter - a moment after the typing pauses, as the search box's is
        // (ApplyReplaceFindNowForTests at once).
        // ###########################################################################################
        private DispatcherTimer? thisReplaceFindTimer;

        private void OnReplaceFindTextChanged(object? sender, TextChangedEventArgs e)
        {
            if (!this.ReplaceBar.IsVisible)
            {
                return;
            }

            if (this.thisReplaceFindTimer is null)
            {
                this.thisReplaceFindTimer = new DispatcherTimer { Interval = BoardTableEditor.SearchDelay };
                this.thisReplaceFindTimer.Tick += (_, _) => this.ApplyReplaceFind();
            }

            this.thisReplaceFindTimer.Stop();
            this.thisReplaceFindTimer.Start();
        }

        private void ApplyReplaceFind(bool stayOnSheet = false)
        {
            this.thisReplaceFindTimer?.Stop();
            this.ApplySearch(this.ReplaceFindBox.Text, literal: true, stayOnSheet);
            this.UpdateReplaceCount();
        }

        internal void ApplyReplaceFindNowForTests() => this.ApplyReplaceFind();

        // "Replace with" typed: the cells draw it at once - no row shows or hides for it.
        private void OnReplaceWithTextChanged(object? sender, TextChangedEventArgs e)
        {
            if (this.ReplaceBar.IsVisible)
            {
                this.UpdateSearchMarks();
                this.UpdateReplaceCount();
            }
        }

        // ###########################################################################################
        // Which rows "Replace all" replaces in - the rows the table shows, the pills and "Find" as they are - and
        // which columns: of a sheet shown in two, the part shown, rows and columns (code review, 2026-11-04: Replace on
        // "Component: Pictures" rewrote the oscilloscope baselines too, unseen, and hidden cells of the pictures). On
        // the sheet shown only (owner request, 2026-11-07: 'remove the checkbox for "On every sheet"').
        // ###########################################################################################
        private Func<BoardTableRow, bool> RowsReplaced()
        {
            BoardTableRowKinds kinds = this.thisFilter;
            BoardTableSearch search = this.thisSearch;
            BoardTableRow? inHand = this.thisRowInHand;
            BoardTableSheetPart part = this.thisPart;

            return row => (ReferenceEquals(row, inHand) || BoardTableSheetParts.Shows(part, row)) && BoardTableRowFilter.Shows(kinds, search, inHand, row, part);
        }

        private IReadOnlyList<int> ColumnsReplaced(BoardTableSheet sheet) => this.ColumnsTickedFor(sheet, this.PartReplacedIn(sheet));

        // ###########################################################################################
        // *** "COLUMNS": WHICH OF THE COLUMNS SHOWN "FIND" LOOKS IN AND "REPLACE ALL" REPLACES IN (owner request,
        // 2026-11-07: "I need some possibility to restrict replacement to specific columns, and not necessarily all";
        // asked, "Yes, let us try this approach"). *** A tick per column the sheet - or its part - shows, every one at
        // first; what is unticked is kept per sheet and part while CRT runs (by the column's NAME), so going back to a
        // sheet finds its choice. The rows shown, the marks and the count all follow it (BoardTableSearch.Literal's
        // columnsOf), so what is seen is what will be replaced.
        // ###########################################################################################
        private readonly Dictionary<string, HashSet<string>> thisColumnsLeftOut = new(StringComparer.Ordinal);

        private static string ColumnsKey(BoardTableSheet sheet, BoardTableSheetPart part) => sheet.Name + "|" + part;

        // The part a sheet is replaced in: the one shown for the sheet shown, a sheet not shown whole.
        private BoardTableSheetPart PartReplacedIn(BoardTableSheet sheet) =>
            ReferenceEquals(sheet, this.thisCurrentSheet) ? this.thisPart : BoardTableSheetPart.Whole;

        private List<int> ColumnsTickedFor(BoardTableSheet sheet, BoardTableSheetPart part)
        {
            List<int> shown = BoardTableSheetParts.ColumnsShown(part, sheet);

            return this.thisColumnsLeftOut.TryGetValue(BoardTableEditor.ColumnsKey(sheet, part), out HashSet<string>? leftOut)
                ? shown.Where(index => !leftOut.Contains(sheet.Columns[index])).ToList()
                : shown;
        }

        // ###########################################################################################
        // The columns of every sheet the replace bar's "Find" looks in, as ticked NOW - null for a sheet with every column
        // ticked. Taken once per search, so the marks drawn later agree with the rows it showed.
        // ###########################################################################################
        private Func<BoardTableSheet, IReadOnlyCollection<int>?> ColumnsFindLooksIn(BoardTableDocument document)
        {
            var columns = new Dictionary<BoardTableSheet, IReadOnlyCollection<int>?>(ReferenceEqualityComparer.Instance);

            foreach (BoardTableSheet sheet in document.Sheets)
            {
                BoardTableSheetPart part = this.PartReplacedIn(sheet);
                columns[sheet] = this.thisColumnsLeftOut.ContainsKey(BoardTableEditor.ColumnsKey(sheet, part))
                    ? this.ColumnsTickedFor(sheet, part).ToHashSet()
                    : null;
            }

            return sheet => columns.GetValueOrDefault(sheet);
        }

        // The ticks, made for the sheet and part shown as the list opens.
        private void OnReplaceColumnsOpening(object? sender, EventArgs e) => this.ShowReplaceColumns();

        // The sheet and part the ticks were last made for - a tick acts on the sheet and part shown, so the list open as
        // another is shown is made again for that one (UpdateBlockEditButtons; code review, 2026-11-08).
        private (BoardTableSheet? Sheet, BoardTableSheetPart Part) thisReplaceColumnsMadeFor;

        private void ShowReplaceColumns()
        {
            this.ReplaceColumnsPanel.Children.Clear();
            this.thisReplaceColumnsMadeFor = (this.thisCurrentSheet, this.thisPart);

            if (this.thisCurrentSheet is not { } sheet)
            {
                return;
            }

            List<int> ticked = this.ColumnsTickedFor(sheet, this.thisPart);

            foreach (int index in BoardTableSheetParts.ColumnsShown(this.thisPart, sheet))
            {
                string name = sheet.Columns[index];
                var box = new CheckBox { Content = name, IsChecked = ticked.Contains(index), Tag = name };
                box.IsCheckedChanged += (_, _) => this.TickReplaceColumn(name, box.IsChecked == true);
                this.ReplaceColumnsPanel.Children.Add(box);
            }
        }

        // One column ticked or unticked: "Find" looked for again, in what is ticked now.
        internal void TickReplaceColumn(string column, bool ticked)
        {
            if (this.thisCurrentSheet is not { } sheet)
            {
                return;
            }

            string key = BoardTableEditor.ColumnsKey(sheet, this.thisPart);

            if (!this.thisColumnsLeftOut.TryGetValue(key, out HashSet<string>? leftOut))
            {
                leftOut = new HashSet<string>(StringComparer.Ordinal);
                this.thisColumnsLeftOut[key] = leftOut;
            }

            bool changed = ticked ? leftOut.Remove(column) : leftOut.Add(column);

            if (leftOut.Count == 0)
            {
                this.thisColumnsLeftOut.Remove(key);
            }

            if (changed)
            {
                this.AfterReplaceColumnsChanged();
            }
        }

        private void OnReplaceEveryColumnClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisCurrentSheet is { } sheet && this.thisColumnsLeftOut.Remove(BoardTableEditor.ColumnsKey(sheet, this.thisPart)))
            {
                this.ShowReplaceColumns();
                this.AfterReplaceColumnsChanged();
            }
        }

        // ###########################################################################################
        // "Unselect all" (owner request, 2026-11-07: "I also need a button for "Unselect all" so I can choose only one
        // column"): every column shown left out, so the one ticked next is the only one looked in.
        // ###########################################################################################
        private void OnReplaceNoColumnClick(object? sender, RoutedEventArgs e) => this.UntickEveryReplaceColumn();

        internal void UntickEveryReplaceColumn()
        {
            if (this.thisCurrentSheet is not { } sheet)
            {
                return;
            }

            var every = BoardTableSheetParts.ColumnsShown(this.thisPart, sheet)
                .Select(index => sheet.Columns[index])
                .ToHashSet(StringComparer.Ordinal);

            string key = BoardTableEditor.ColumnsKey(sheet, this.thisPart);

            if (this.thisColumnsLeftOut.TryGetValue(key, out HashSet<string>? leftOut) && leftOut.SetEquals(every))
            {
                return;
            }

            this.thisColumnsLeftOut[key] = every;
            this.ShowReplaceColumns();
            this.AfterReplaceColumnsChanged();
        }

        // ###########################################################################################
        // "Find" looked for again in the columns ticked now - on THIS sheet and part, whatever they show (code review,
        // 2026-11-08). The ticks are theirs alone: with every column unticked, "Find" shows none of the sheet's rows, and
        // moving to another sheet holding the text - as typing in "Find" does - left the list open acting on that other
        // sheet, so the one column wanted could never be ticked.
        // ###########################################################################################
        private void AfterReplaceColumnsChanged()
        {
            if (this.ReplaceBar.IsVisible)
            {
                this.ApplyReplaceFind(stayOnSheet: true);
            }

            this.ShowReplaceColumnsLabel();
        }

        // "Columns: All", "Columns: 2 of 7" - for the sheet and part shown.
        private void ShowReplaceColumnsLabel()
        {
            if (this.thisCurrentSheet is not { } sheet)
            {
                this.ReplaceColumnsButton.Content = BoardTableEditWording.ColumnsChoice(0, 0);
                return;
            }

            this.ReplaceColumnsButton.Content = BoardTableEditWording.ColumnsChoice(
                this.ColumnsTickedFor(sheet, this.thisPart).Count,
                BoardTableSheetParts.ColumnsShown(this.thisPart, sheet).Count);
        }

        // The label on "Columns" - for the tests.
        internal string ReplaceColumnsLabelForTests => this.ReplaceColumnsButton.Content as string ?? string.Empty;

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

            if (this.ColumnsReplaced(sheet).Count == 0)
            {
                this.ShowStatus(BoardTableEditWording.NoColumnsTicked);
                return;
            }

            // What "Find" holds now is the filter it replaces under - typing that has not paused yet included - on THIS
            // sheet and part (code review, 2026-11-08): looked for as typing does, it moved the table to another sheet holding
            // the text, while the replace went on in this one, taken whole - its unticked columns rewritten out of sight.
            if (!string.Equals(find, this.thisSearchText, StringComparison.Ordinal) || !this.thisSearchIsLiteral)
            {
                this.ApplyReplaceFind(stayOnSheet: true);
            }

            BoardTableReplaceResult result = document.ReplaceAll(find, this.ReplaceWithBox.Text, [sheet], this.RowsReplaced(), this.ColumnsReplaced);

            // The rows replaced in stay in sight, as a search decides its rows once: what they read now is seen.
            this.ShowStatus(BoardTableEditWording.Replaced(result, find));
            this.UpdateSheetTabs();
            this.UpdateToolbar();
            this.UpdateSearchMarks();
        }

        // ###########################################################################################
        // The count on "Replace all": how many places it will replace on the sheet shown (owner request, 2026-11-07: "a
        // counter badge for how many places it will replace") - BoardTableDocument.CountReplacements, by ReplaceAll's
        // own rules.
        //
        // *** WORKED OUT AGAIN ONLY WHEN SOMETHING IT IS WORKED OUT FROM HAS MOVED (code review, 2026-11-08). *** It follows
        // every toolbar update - a cell or the selection moved - and each count walks every cell ticked on the sheet, tens
        // of thousands of string operations on a sheet of 5,000 rows, on the UI thread at every arrow key. So it is kept
        // by what it depends on (ReplaceCountKey): the sheet and part, the pills and "Find" as applied, the row in hand
        // only where nothing else would show it (RowsReplaced), the texts, the columns ticked - and the table's content,
        // whose every change moves thisContentRevision (OnCellEdited, OnDocumentChanged).
        // ###########################################################################################
        private readonly record struct ReplaceCountKey(
            BoardTableDocument Document,
            BoardTableSheet Sheet,
            BoardTableSheetPart Part,
            BoardTableRowKinds Filter,
            BoardTableSearch Search,
            BoardTableRow? RowInHandOnly,
            string Find,
            string With,
            string Columns,
            int Revision);

        private ReplaceCountKey? thisReplaceCountFor;

        // Moved by every change of the table's content - a cell typed in, rows refreshed, an undo.
        private int thisContentRevision;

        // How often the count was worked out - for the tests.
        internal int ReplaceCountsForTests { get; private set; }

        private void UpdateReplaceCount()
        {
            if (!this.ReplaceBar.IsVisible)
            {
                return;
            }

            if (this.thisDocument is not { } document || this.thisCurrentSheet is not { } sheet)
            {
                this.thisReplaceCountFor = null;
                this.ShowReplaceCount(0);
                return;
            }

            var key = new ReplaceCountKey(
                document,
                sheet,
                this.thisPart,
                this.thisFilter,
                this.thisSearch,
                this.RowInHandOnly(),
                this.ReplaceFindBox.Text ?? string.Empty,
                this.ReplaceWithBox.Text ?? string.Empty,
                string.Join(",", this.ColumnsReplaced(sheet)),
                this.thisContentRevision);

            if (this.thisReplaceCountFor == key)
            {
                return;
            }

            this.thisReplaceCountFor = key;
            this.ReplaceCountsForTests++;
            this.ShowReplaceCount(document.CountReplacements(key.Find, key.With, [sheet], this.RowsReplaced(), this.ColumnsReplaced));
        }

        // The row in hand when only being in hand shows it - the one row whose place in the count the cursor decides.
        private BoardTableRow? RowInHandOnly() =>
            this.thisRowInHand is { } inHand &&
            !(BoardTableSheetParts.Shows(this.thisPart, inHand) && BoardTableRowFilter.Shows(this.thisFilter, this.thisSearch, null, inHand, this.thisPart))
                ? inHand
                : null;

        private void ShowReplaceCount(int count)
        {
            this.ReplaceCountText.Text = count.ToString("N0", CultureInfo.InvariantCulture);
            ToolTip.SetTip(this.ReplaceAllButton, count == 1 ? "Replaces 1 place on this sheet" : $"Replaces {count.ToString("N0", CultureInfo.InvariantCulture)} places on this sheet");
        }

        // The count "Replace all" shows - for the tests.
        internal string ReplaceCountForTests => this.ReplaceCountText.Text ?? string.Empty;

        private void WireBlockEdits()
        {
            this.ReplaceFindBox.KeyDown += this.OnReplaceBoxKeyDown;
            this.ReplaceWithBox.KeyDown += this.OnReplaceBoxKeyDown;
            this.ReplaceFindBox.TextChanged += this.OnReplaceFindTextChanged;
            this.ReplaceWithBox.TextChanged += this.OnReplaceWithTextChanged;

            if (this.ReplaceColumnsButton.Flyout is { } columns)
            {
                columns.Opened += this.OnReplaceColumnsOpening;
            }

            ToolTip.SetTip(this.ReplaceColumnsButton, BoardTableEditWording.ColumnsChoiceTip);
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
            this.ReplaceButton.IsVisible = !this.thisIsReadOnly;
            this.ReplaceButton.IsEnabled = this.thisCurrentSheet is not null;

            if (this.thisIsReadOnly)
            {
                this.LetGoOfReplaceBar();
            }

            // The count follows whatever the toolbar follows - a sheet shown, a pill picked, a change made - and so
            // does the "Columns" label, of the sheet shown, and the list's ticks while it is open.
            this.UpdateReplaceCount();
            this.ShowReplaceColumnsLabel();

            if (this.ReplaceColumnsButton.Flyout is { IsOpen: true } &&
                (!ReferenceEquals(this.thisReplaceColumnsMadeFor.Sheet, this.thisCurrentSheet) || this.thisReplaceColumnsMadeFor.Part != this.thisPart))
            {
                this.ShowReplaceColumns();
            }
        }

        // ###########################################################################################
        // A key with the platform's command modifier alone - Ctrl, or Cmd on macOS - for the shortcuts
        // the platform's hotkey list has no entry for (find, replace). Ctrl where the
        // platform does not say, as on the headless one the tests run on.
        // ###########################################################################################
        private static bool IsCommand(KeyEventArgs e, Key key)
        {
            KeyModifiers command = Application.Current?.PlatformSettings?.HotkeyConfiguration.CommandModifiers ?? KeyModifiers.Control;

            return e.Key == key && e.KeyModifiers == command;
        }
    }
}
