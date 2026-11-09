using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE ROWS OF THE COMPONENTS PICKED (owner request, 2026-11-08: "When in "Table" view I think it should filter the
    // selected components so it becomes easy to overview those in the different menus - also, the "Replace" will only work
    // on those filtered rows"). A host that picks components elsewhere - the Contribute tab's list - hands them over
    // (ShowOnlyComponents): on a sheet with a "Board label" column only their rows show, one or several; a sheet of the
    // board as a whole shows every row. It narrows the table as a pill or a search does, so "Replace" replaces in these
    // rows alone, a paste adds no rows and no row moves - and the line over the table says whose rows these are. The row the
    // cursor is on stays until the cursor leaves it, as under a pill.
    //
    // Part of the BoardTableEditor partial class - see BoardTableEditor.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // The labels whose rows a component sheet shows - null for every component's.
        private HashSet<string>? thisComponentsShown;

        // The same, as the host handed them over - its list's order - for the line over the table.
        private List<string> thisComponentsShownNamed = [];

        internal IReadOnlyCollection<string> ComponentsShownForTests => this.thisComponentsShown ?? [];

        // ###########################################################################################
        // Only the rows of `labels` on the component sheets - none for every component's. Nothing done when it is what is
        // shown already.
        // ###########################################################################################
        internal void ShowOnlyComponents(IReadOnlyCollection<string> labels)
        {
            ArgumentNullException.ThrowIfNull(labels);

            HashSet<string>? wanted = labels.Count == 0
                ? null
                : new HashSet<string>(labels.Select(label => label.Trim()).Where(label => label.Length > 0), StringComparer.OrdinalIgnoreCase);

            if (wanted is { Count: 0 })
            {
                wanted = null;
            }

            if (wanted is null ? this.thisComponentsShown is null : this.thisComponentsShown?.SetEquals(wanted) == true)
            {
                return;
            }

            this.thisComponentsShown = wanted;
            this.thisComponentsShownNamed = wanted is null ? [] : labels.Select(label => label.Trim()).Where(wanted.Contains).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            this.TableGrid.CommitEdit();
            this.UpdateRowsDraggable();
            this.ApplyViewFilter();
            this.ShowComponentsShownLine();

            // Under a pill or a search, which sheets show anything moves with them (ShowsRowsOfPicked) - the tabs, or the host's
            // entries. The sheet shown stays: it is the one just asked for.
            this.UpdateSheetTabs();
            this.UpdateToolbar();
        }

        // ###########################################################################################
        // Whether the sheet shown shows a row for the components picked: every row of a sheet with no "Board label", and of
        // one, a row of one of them - or a row still blank, being filled in. The column looked up ONCE for a whole pass (code
        // review, 2026-11-09: once per row, a new list of the sheet's columns each time).
        // ###########################################################################################
        private Func<BoardTableRow, bool> ShowsComponentOf()
        {
            if (this.thisComponentsShown is not { } labels || this.thisCurrentSheet is not { } sheet)
            {
                return _ => true;
            }

            int column = BoardTableEditor.LabelColumnOf(sheet);

            return column < 0
                ? _ => true
                : row => row.State is BoardTableRowState.Blank || column >= row.Cells.Count || labels.Contains(row.Cells[column].Text.Trim());
        }

        private static int LabelColumnOf(BoardTableSheet sheet)
        {
            int index = 0;

            foreach (string column in sheet.Columns)
            {
                if (string.Equals(column, BoardWorkbookSchema.ColBoardLabel, StringComparison.Ordinal))
                {
                    return index;
                }

                index++;
            }

            return -1;
        }

        // Whether the sheet shown is narrowed to the components picked - only a sheet with a "Board label" column is.
        private bool IsNarrowedToComponents =>
            this.thisComponentsShown is not null &&
            this.thisCurrentSheet is { } sheet &&
            BoardTableEditor.LabelColumnOf(sheet) >= 0;

        // ###########################################################################################
        // Whether `sheet` (its `part`) shows any row under `filter` and `search` once narrowed to the components picked - every
        // sheet with no "Board label", or with none picked, as far as they go (third code review, 2026-11-09: the entries and the
        // sheet a pill moved to went by the pills alone, and showed rows the narrowing then hid).
        // ###########################################################################################
        private bool ShowsRowsOfPicked(BoardTableSheet sheet, BoardTableSheetPart part, BoardTableRowKinds filter, BoardTableSearch search)
        {
            if (this.thisComponentsShown is not { } labels)
            {
                return true;
            }

            int column = BoardTableEditor.LabelColumnOf(sheet);

            if (column < 0)
            {
                return true;
            }

            return BoardTableRowFilter.RowsShown(sheet.Rows, filter, search, rowInHand: null, part)
                .Any(row => row.State is BoardTableRowState.Blank || (column < row.Cells.Count && labels.Contains(row.Cells[column].Text.Trim())));
        }

        // The rows shown on the sheet shown, narrowed to the components picked - the row in hand kept.
        private List<BoardTableRow> NarrowedToComponents(List<BoardTableRow> rows, BoardTableRow? rowInHand)
        {
            if (this.thisComponentsShown is null)
            {
                return rows;
            }

            Func<BoardTableRow, bool> ofPicked = this.ShowsComponentOf();

            return rows.Where(row => ReferenceEquals(row, rowInHand) || ofPicked(row)).ToList();
        }

        // What the replace count depends on of this.
        private string ComponentsShownKey => string.Join(",", this.thisComponentsShownNamed);

        // The line over the table: whose rows these are - on a component sheet only.
        private void ShowComponentsShownLine()
        {
            bool shows = this.IsNarrowedToComponents;

            this.ComponentsShownText.IsVisible = shows;
            this.ComponentsShownText.Text = shows
                ? BoardTableEditWording.ComponentsShown(this.thisComponentsShownNamed)
                : string.Empty;
        }
    }
}
