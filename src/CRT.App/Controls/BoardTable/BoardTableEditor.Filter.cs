using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Data.Converters;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE TABLE EDITOR'S COLOUR KEY AS THE FILTER - which kinds of row the picked pills show, the
    // user's own pick apart from what the grid shows, and the rows the grid shows. Split out of
    // BoardTableEditor.axaml.cs (code review, 2026-10-04: the main part had passed the ~1,500-line
    // mark). The file map is in BoardTableEditor.axaml.cs.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // The kinds of row the colour key's picked pills show - None shows every row
        // (BoardTableRowFilter; owner request, 2026-10-02 - it replaced "Show changes only").
        private BoardTableRowKinds thisFilter;

        // What the user PICKED. The filter itself is off for a table where it would show nothing
        // at all (no row of those kinds anywhere - every table with nothing published, for the
        // change kinds) - and comes back on for the next table that has such rows, rather than
        // staying off because one table without any was looked at in between (2026-09-26).
        private BoardTableRowKinds thisFilterWanted;

        // ###########################################################################################
        // *** WHAT THE GRID SHOWS IS A LIST OF THE TABLE'S OWN, AND THE VIEW ON IT HAS NO FILTER (code
        // review, 2026-11-02). *** A filtered DataGridCollectionView takes the row out and puts it back
        // at every row commit, which leaves the grid with no current cell and no selection: Enter on
        // the last row shown, Ctrl+D, Ctrl+F, Ctrl+V with a row's edit open all lost the cursor, and
        // the next text typed went nowhere or, the grid putting the cursor on the first row shown, over
        // another row's cell. Repairing that after each commit was a repair per way of committing. With
        // no filter a commit touches no list - as in a table with nothing picked - so there is nothing
        // to repair: this list is sheet.Rows less what the pills and the search hide, changed one row
        // at a time as rows come and go (ShowRows), and never under a row being edited.
        //
        // *** Do not put a Filter back on the view. *** BoardTableEditorNarrowedTests holds it null.
        // ###########################################################################################
        private ShownRowList? thisShownRows;

        // The rows the grid shows now - for the tests.
        internal IReadOnlyList<BoardTableRow> RowsShownForTests => this.thisShownRows?.ToList() ?? [];

        // ###########################################################################################
        // The row the cursor is on, which stays shown until the cursor leaves it, whatever is typed in
        // it (BoardTableRowFilter's header). A pick or a search keeps nothing (ApplyViewFilter).
        // ###########################################################################################
        private BoardTableRow? thisRowInHand;

        private bool thisShowRowsPosted;

        private bool thisRowsChangedTogether;

        // ###########################################################################################
        // *** THE COLOUR KEY IS THE FILTER (owner request, 2026-10-02: the counts clickable, "and
        // even do remove the "Show changes only" so it works in the same unified way"). *** Each
        // pill is a kind of row; the picked ones show only rows of any of them, and hide the tabs
        // of sheets with none (BoardTableRowFilter, BoardTableDocument.SheetsShown). Moving rows is
        // off meanwhile - a position among rows that cannot be seen means nothing.
        //
        // Filter is what the grid shows now; FilterWanted is what the user picked. They differ only
        // for a table where the picked kinds show nothing at all, which opens unfiltered (Attach).
        // ###########################################################################################
        internal BoardTableRowKinds Filter
        {
            get => this.thisFilter;
            set
            {
                bool changed = this.thisFilterWanted != value;

                this.thisFilterWanted = value;
                this.ApplyFilter(value);

                if (changed)
                    this.FilterWantedChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        // The user's own pick, which the filter follows wherever it shows something - what a host
        // remembers between runs (CRT's Maintainer tab).
        internal BoardTableRowKinds FilterWanted => this.thisFilterWanted;

        // ###########################################################################################
        // The user's PICK changed - so a host can remember it (CRT's Maintainer tab does). Raised
        // HERE ONLY, where the pick changes: Attach and ApplyFilter change what the grid SHOWS
        // alone. Raising it from those would write "nothing picked" back as the user's choice every
        // time a table with no such rows opened unfiltered. The Contribute tab does not listen.
        // ###########################################################################################
        internal event EventHandler? FilterWantedChanged;

        // ###########################################################################################
        // A pick made in ANOTHER table that shows the same thing - the Maintainer tab's submission
        // table and its board table (code review, 2026-10-04: each wrote the one remembered pick
        // without telling the other, so the next launch opened both on whichever was picked last).
        // Taken as this table's pick and applied the way a table being opened applies one (Attach):
        // not at all where this table would show nothing of it. Raises nothing - the host has
        // remembered it already, and the other table told it.
        // ###########################################################################################
        internal void UseFilterWanted(BoardTableRowKinds kinds)
        {
            if (this.thisFilterWanted == kinds)
            {
                return;
            }

            this.thisFilterWanted = kinds;

            this.ApplyFilter(this.thisDocument is { } document && !document.HasRowsShownBy(kinds)
                ? BoardTableRowKinds.None
                : kinds);
        }

        // One pill clicked: its kind picked, or put back. A pill counting nothing cannot be picked
        // (owner request, 2026-10-03 - BoardTableDocument.CanToggle); its style drops the hand cursor.
        internal void TogglePill(BoardTableRowKinds kind)
        {
            if (this.thisDocument is not null && !this.thisDocument.CanToggle(this.thisFilter, kind))
            {
                return;
            }

            this.Filter = this.thisFilter ^ kind;
        }

        private void OnLegendPillTapped(object? sender, TappedEventArgs e)
        {
            BoardTableRowKinds kind = this.PillKindOf(sender);

            if (kind == BoardTableRowKinds.None)
            {
                return;
            }

            e.Handled = true;
            this.TogglePill(kind);
            this.FocusGrid();
        }

        // ###########################################################################################
        // Enter or Space on a pill reached with Tab does what a click does (code review, 2026-10-13).
        // The focus stays on the pill, so the next pill is one Tab away - a click hands it to the
        // grid instead, where the mouse user goes next.
        // ###########################################################################################
        private void OnLegendPillKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key is not (Key.Enter or Key.Space))
            {
                return;
            }

            BoardTableRowKinds kind = this.PillKindOf(sender);

            if (kind == BoardTableRowKinds.None)
            {
                return;
            }

            e.Handled = true;
            this.TogglePill(kind);
        }

        private BoardTableRowKinds PillKindOf(object? sender) => sender switch
        {
            Border border when ReferenceEquals(border, this.AddedPill) => BoardTableRowKinds.Added,
            Border border when ReferenceEquals(border, this.ModifiedPill) => BoardTableRowKinds.Modified,
            Border border when ReferenceEquals(border, this.DeletedPill) => BoardTableRowKinds.Deleted,
            Border border when ReferenceEquals(border, this.ErrorsPill) => BoardTableRowKinds.Errors,
            Border border when ReferenceEquals(border, this.WarningsPill) => BoardTableRowKinds.Warnings,
            _ => BoardTableRowKinds.None
        };

        // ###########################################################################################
        // Shows `kinds` without touching the user's pick (thisFilterWanted).
        //
        // *** IT ALSO HIDES THE TABS OF SHEETS IT WOULD SHOW NOTHING OF *** (owner request,
        // 2026-09-26) - BoardTableDocument.SheetsShown. Picked while on such a sheet, the table
        // moves to the first sheet that still has a tab: picking "Errors" goes to the errors.
        // ###########################################################################################
        private void ApplyFilter(BoardTableRowKinds kinds)
        {
            if (this.thisFilter == kinds)
            {
                return;
            }

            this.thisFilter = kinds;
            this.TableGrid.CommitEdit();
            this.UpdateRowsDraggable();

            this.ApplyViewFilter();
            this.ShowASheetWithRowsShown();
            this.UpdateSheetTabs();
            this.UpdateToolbar();
            this.UpdateSearchMarks();
        }

        // Picked, or searched, while on a sheet that shows nothing of it: the table moves to the
        // first sheet that does. Only a sheet of the document on screen - Attach applies the filter
        // before the new document's sheet is chosen. A search finding nothing anywhere stays on the
        // sheet on screen (SheetToShow hands it back), so it is not chosen again.
        private void ShowASheetWithRowsShown()
        {
            if (this.IsNarrowed &&
                this.thisDocument is not null &&
                this.thisCurrentSheet is not null &&
                this.thisDocument.Sheets.Contains(this.thisCurrentSheet) &&
                !this.thisDocument.SheetsShown(this.thisFilter, this.thisSearch, current: null).Contains(this.thisCurrentSheet))
            {
                BoardTableSheet sheet = this.thisDocument.SheetToShow(this.thisCurrentSheet.Name, this.thisFilter, this.thisSearch);

                if (!ReferenceEquals(sheet, this.thisCurrentSheet))
                {
                    this.SelectSheet(sheet);
                }
            }

            this.ShowAPartWithRowsShown();
        }

        // ###########################################################################################
        // What the pills and the search show has changed: a new question, so the row in hand is kept
        // only if it is an answer - picking "Errors" hides the row the cursor is on when it has none.
        // All at once: most rows may come or go, and one at a time a pill took 1-4 seconds on a sheet of
        // 5,000 rows where a Reset takes about 0.2, as a Filter set on the view did (measured, 2026-11-02).
        //
        // *** THE GRID MUST BE TOLD THE VIEW'S FILTER IS OURS. *** Its column-filtering model
        // (FilteringModel) owns a view's Filter by default and writes its own - empty - predicate
        // over it whenever it takes the view or is attached again. It lost "Show changes only", with
        // the box still ticked, on every sheet switch and every switch of the MAIN tabs; today it
        // would put back a filter, and with it the commit that takes the cursor away (thisShownRows).
        // OwnsViewFilter = false makes it leave the view alone; the table uses none of its filtering.
        // ###########################################################################################
        private void ApplyViewFilter()
        {
            if (this.TableGrid.FilteringModel is { } model)
            {
                model.OwnsViewFilter = false;
            }

            this.thisRowInHand = null;

            if (this.thisShownRows is { } shown && this.thisCurrentSheet is { } sheet)
            {
                List<BoardTableRow> rows = BoardTableRowFilter.RowsShown(sheet.Rows, this.thisFilter, this.thisSearch, rowInHand: null, this.thisPart);

                if (!rows.SequenceEqual(shown))
                {
                    shown.ReplaceAll(rows);
                }
            }

            this.thisRowInHand = this.CurrentCell?.Row;
        }

        // ###########################################################################################
        // The rows the grid shows brought up to date, keeping the row the cursor is on now.
        //
        // The cursor on a CELL (CurrentCell, not CurrentRow): the grid sits on the first row's marker
        // column by itself, in a table nobody has touched, and that is no row anybody is working in -
        // its last change undone from elsewhere, it goes. So the marker column taken alone says nothing:
        // the row in hand stays what it was, also when that row's own marker column is clicked.
        // ###########################################################################################
        private void ShowRows()
        {
            if (this.CurrentCell is { } cell)
            {
                this.thisRowInHand = cell.Row;
            }

            this.SyncShownRows();
        }

        // ###########################################################################################
        // One row at a time, so what stays keeps its place, its scroll position and the cursor. With
        // nothing picked that is all of them.
        //
        // Not under a row being edited while narrowed: removing another row resets the grid's current
        // cell around it. The row's edit ending comes back here (WireShownRows).
        // ###########################################################################################
        private void SyncShownRows()
        {
            if (this.thisShownRows is not { } shown || this.thisCurrentSheet is not { } sheet)
            {
                return;
            }

            if (this.IsNarrowed && this.thisView?.IsEditingItem == true)
            {
                return;
            }

            this.RowsShownAskedForTests++;

            BoardTableSheet.SyncRows(shown, BoardTableRowFilter.RowsShown(sheet.Rows, this.thisFilter, this.thisSearch, this.thisRowInHand, this.thisPart));
        }

        // How many times the sheet has been gone over for the rows to show - for tests.
        internal int RowsShownAskedForTests { get; private set; }

        // Soon, once the grid has finished what it is doing - moving its cursor, or ending an edit.
        private void PostShowRows()
        {
            if (!this.IsNarrowed || this.thisShowRowsPosted)
            {
                return;
            }

            this.thisShowRowsPosted = true;

            Dispatcher.UIThread.Post(() =>
            {
                this.thisShowRowsPosted = false;
                this.ShowRows();
            }, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // The cursor moved, so the row it left goes if only the cursor kept it. Both events, and the
        // answer read AFTER them (measured with a probe, 2026-11-02): a click raises SelectedCellsChanged
        // with the grid's cursor still on the row before, and no CurrentCellChanged at all.
        // ###########################################################################################
        private void WireShownRows()
        {
            this.TableGrid.CurrentCellChanged += (_, _) => this.PostShowRows();
            this.TableGrid.SelectedCellsChanged += (_, _) => this.PostShowRows();
            this.TableGrid.RowEditEnded += (_, _) => this.PostShowRows();
        }

        // `row`, on a table just built again, stays shown as if the cursor had never left it.
        private void KeepInHand(BoardTableRow row)
        {
            this.thisRowInHand = row;
            this.SyncShownRows();
        }

        // ###########################################################################################
        // A row of the sheet added, removed or moved. With nothing picked the list IS the sheet's, so
        // the change is made as the sheet made it: the grid sees what it would on the sheet's own rows.
        // ###########################################################################################
        private void OnSheetRowsChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (this.IsNarrowed || this.thisShownRows is not { } shown)
            {
                // Many rows changing as one are gone over once, as they end (code review, 2026-11-03): per
                // row, deleting thousands under a pill cost selected rows x sheet rows on the UI thread.
                if (this.thisCurrentSheet?.RowsAreChangingTogether == true)
                {
                    this.thisRowsChangedTogether = true;
                    return;
                }

                this.ShowRows();
                return;
            }

            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    shown.Insert(e.NewStartingIndex, (BoardTableRow)e.NewItems![0]!);
                    break;

                case NotifyCollectionChangedAction.Remove:
                    shown.RemoveAt(e.OldStartingIndex);
                    break;

                case NotifyCollectionChangedAction.Move:
                    shown.Move(e.OldStartingIndex, e.NewStartingIndex);
                    break;

                case NotifyCollectionChangedAction.Replace:
                    shown[e.NewStartingIndex] = (BoardTableRow)e.NewItems![0]!;
                    break;

                default:
                    this.SyncShownRows();
                    break;
            }
        }

        private void OnSheetRowsChangedTogether(object? sender, EventArgs e)
        {
            if (this.thisRowsChangedTogether)
            {
                this.thisRowsChangedTogether = false;
                this.ShowRows();
            }
        }

        // The rows the grid shows: changed row by row, or all at once with one Reset (ApplyViewFilter).
        private sealed class ShownRowList(IEnumerable<BoardTableRow> rows) : ObservableCollection<BoardTableRow>(rows)
        {
            public void ReplaceAll(IEnumerable<BoardTableRow> rows)
            {
                this.CheckReentrancy();
                this.Items.Clear();

                foreach (BoardTableRow row in rows)
                {
                    this.Items.Add(row);
                }

                this.OnPropertyChanged(new PropertyChangedEventArgs(nameof(this.Count)));
                this.OnPropertyChanged(new PropertyChangedEventArgs("Item[]"));
                this.OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
            }
        }
    }
}
