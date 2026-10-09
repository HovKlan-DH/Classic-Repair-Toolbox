using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Data.Converters;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Styling;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE TABLE EDITOR'S COLOUR KEY AS THE FILTER - which kinds of row the picked pills show, the
    // user's own pick apart from what the grid shows, and the view's filter itself. Split out of
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
        // time a table with no such rows opened unfiltered. The Drafts tab does not listen.
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
            BoardTableRowKinds kind = sender switch
            {
                Border border when ReferenceEquals(border, this.AddedPill) => BoardTableRowKinds.Added,
                Border border when ReferenceEquals(border, this.ModifiedPill) => BoardTableRowKinds.Modified,
                Border border when ReferenceEquals(border, this.DeletedPill) => BoardTableRowKinds.Deleted,
                Border border when ReferenceEquals(border, this.ErrorsPill) => BoardTableRowKinds.Errors,
                Border border when ReferenceEquals(border, this.WarningsPill) => BoardTableRowKinds.Warnings,
                _ => BoardTableRowKinds.None
            };

            if (kind == BoardTableRowKinds.None)
            {
                return;
            }

            e.Handled = true;
            this.TogglePill(kind);
            this.FocusGrid();
        }

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
        }

        // ###########################################################################################
        // *** THE GRID MUST BE TOLD THE VIEW'S FILTER IS OURS. *** Its column-filtering model
        // (FilteringModel) owns a view's Filter by default and writes its own - empty - predicate
        // over it whenever it takes the view or is attached again. So "Show changes only" was lost,
        // with the box still ticked, first on every sheet switch and then on every switch of the
        // MAIN tabs (Drafts -> Contribute -> Drafts), both reported. OwnsViewFilter = false makes
        // it leave the filter alone; the table uses none of its column filtering.
        // ###########################################################################################
        private void ApplyViewFilter()
        {
            if (this.TableGrid.FilteringModel is { } model)
            {
                model.OwnsViewFilter = false;
            }

            if (this.thisView is not null)
            {
                BoardTableRowKinds kinds = this.thisFilter;
                BoardTableSearch search = this.thisSearch;

                // The picked pills and the search box together: a row both show (2026-10-02).
                this.thisView.Filter = kinds == BoardTableRowKinds.None && !search.IsActive
                    ? null
                    : item => item is BoardTableRow row && BoardTableRowFilter.Shows(kinds, row) && search.Shows(row);
            }
        }
    }
}
