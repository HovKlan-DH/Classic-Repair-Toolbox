using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // WHAT PICKS WHAT THE TABLE SHOWS. On the Contribute tab its left column (owner request, 2026-11-03: "when I
    // select e.g. 'Board: Schematic images' then I will see the table for 'Board schematics' in the content
    // area"): its host says so (SheetsPickedElsewhere), so the sheet tabs go and the host shows a sheet by
    // ShowSheetPart. On the Maintainer tab the sheet tabs, under the same names and headings (BoardTableTabs;
    // owner request, 2026-11-05: "group and order into what is logical for the new setup").
    //
    // "Component images" is shown in two either way - the rules are BoardTableSheetParts - which only says which
    // rows and columns the grid is given: the sheet stays one sheet, and everything else about it goes as ever.
    // Rows do not move by their grip in such a part (IsNarrowed), a position among rows that cannot be seen
    // meaning nothing.
    //
    // What moves the table by itself - a pill or a search going to the rows it found, an undo going to its
    // change - is told by SheetShown, so the host's column follows.
    //
    // The file map is in BoardTableEditor.axaml.cs.
    // ###########################################################################################

    // One tab of the strip: what it is, and the sheet of this table it shows.
    internal sealed record BoardTableSheetTab(BoardTableTab Tab, BoardTableSheet Sheet);

    public partial class BoardTableEditor
    {
        private BoardTableSheetPart thisPart = BoardTableSheetPart.Whole;

        // What SheetShown last told.
        private (string Sheet, BoardTableSheetPart Part)? thisSheetToldShown;

        internal BoardTableSheetPart SheetPart => this.thisPart;

        // The sheet shown, or the part of it, is another now - by whatever moved it.
        internal event EventHandler? SheetShown;

        // What the pills and the search leave to show has changed.
        internal event EventHandler? SheetsShownChanged;

        // The host's own column picks the sheets, so no tabs.
        internal bool SheetsPickedElsewhere
        {
            get => !this.SheetTabs.IsVisible;
            set => this.SheetTabs.IsVisible = !value;
        }

        // ###########################################################################################
        // A sheet, or a part of one, shown - and nothing done when it already is, so the cursor stays. A
        // pill or a search leaving nothing of the part asked for shows the other (ChoosePartOf).
        // ###########################################################################################
        internal bool ShowSheetPart(string sheetName, BoardTableSheetPart part)
        {
            if (this.thisDocument?.FindSheet(sheetName) is not BoardTableSheet sheet)
            {
                return false;
            }

            if (ReferenceEquals(sheet, this.thisCurrentSheet) && part == this.thisPart)
            {
                return true;
            }

            this.thisPart = part;
            this.SelectSheet(sheet);

            return true;
        }

        // ###########################################################################################
        // Whether a sheet or part has anything to show under the pills and the search - a tab would be
        // gone. The draft is gone over ONCE however many are asked about (code review, 2026-11-04: each of
        // the left column's entries went over every sheet, on the UI thread after each edit).
        // ###########################################################################################
        internal Func<string, BoardTableSheetPart, bool> WhatShows()
        {
            if (this.thisDocument is not { } document)
            {
                return (_, _) => true;
            }

            BoardTableRowKinds filter = this.thisFilter;
            BoardTableSearch search = this.thisSearch;
            IReadOnlyList<BoardTableSheet> shown = document.SheetsShown(filter, search, current: null);
            bool narrowed = document.HasRowsShownBy(filter, search);

            return (sheetName, part) =>
            {
                if (document.FindSheet(sheetName) is not BoardTableSheet sheet)
                {
                    return true;
                }

                if (ReferenceEquals(sheet, this.thisCurrentSheet) && (part == this.thisPart || part == BoardTableSheetPart.Whole))
                {
                    return true;
                }

                // Only a sheet in two has parts to tell apart. A sheet shown whole goes by the document's answer,
                // which counts a row just inserted (code review, 2026-11-05) - ShowsAnything does not, or it would
                // keep both parts for one such row, and its tab went with the insert looking as if it had failed.
                // Narrowed to the components picked as well, a component sheet showing none of theirs shows nothing
                // (third code review, 2026-11-09: "Errors" left their entries on, for rows the narrowing hid).
                return shown.Contains(sheet) &&
                       (!narrowed || part == BoardTableSheetPart.Whole || BoardTableSheetParts.ShowsAnything(sheet, part, filter, search)) &&
                       (!narrowed || this.ShowsRowsOfPicked(sheet, part, filter, search));
            };
        }

        // ###########################################################################################
        // The first row of a component on the sheet shown, with the cursor on it - a component picked in
        // the list while the table is what shows. False, said on the status line, when it has none here.
        // ###########################################################################################
        internal bool GoToRowsOfComponent(string label)
        {
            if (this.thisCurrentSheet is not { } sheet)
            {
                return false;
            }

            int column = sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel);

            if (column < 0)
            {
                return false;
            }

            BoardTableRow? row = this.ShownRows().FirstOrDefault(candidate =>
                !candidate.IsDeleted &&
                string.Equals(candidate.Cells[column].Text.Trim(), label.Trim(), StringComparison.OrdinalIgnoreCase));

            if (row is null)
            {
                this.ShowStatus(BoardTableEditWording.NoRowsOfComponent(label));
                return false;
            }

            this.ShowStatus(string.Empty);
            this.SelectCell(row, column);

            return true;
        }

        // The columns of the sheet the grid shows, by their index in the sheet.
        private List<int> ColumnsShownNow() =>
            this.thisCurrentSheet is { } sheet ? BoardTableSheetParts.ColumnsShown(this.thisPart, sheet) : [];

        private void ChoosePartOf(BoardTableSheet sheet) =>
            this.thisPart = BoardTableSheetParts.PartToShow(sheet, this.thisPart, this.thisFilter, this.thisSearch);

        // A pill or a search with nothing in this part of the sheet: to the part that has it.
        private void ShowAPartWithRowsShown()
        {
            if (this.thisCurrentSheet is { } sheet &&
                this.thisDocument?.Sheets.Contains(sheet) == true &&
                BoardTableSheetParts.PartToShow(sheet, this.thisPart, this.thisFilter, this.thisSearch) != this.thisPart)
            {
                this.SelectSheet(sheet);
            }
        }

        // ###########################################################################################
        // The row a step goes to, on a sheet in two. Only a row the step really CHANGED moves the view to
        // the part it belongs to; one standing as it stood before (`before`, the sheet as it was) only slid
        // into the step's place - a picture added below the last one and undone leaves the first baseline at
        // its index, as a board file holds the pictures first (code review, 2026-11-04: the view jumped there,
        // and the left column's entry with it). The view then stays, the cursor on the nearest row it shows.
        // ###########################################################################################
        private BoardTableRow RowToGoToAfterStep(BoardTableSheet sheet, BoardTableRow row, BoardTableSheet? before)
        {
            if (BoardTableSheetParts.Shows(this.thisPart, row))
            {
                return row;
            }

            if (before is null || !BoardTableSheetParts.StandsIn(before, row))
            {
                this.thisPart = BoardTableSheetParts.PartOf(row);
                this.SelectSheet(sheet);

                return row;
            }

            return BoardTableSheetParts.NearestShown(sheet, this.thisPart, sheet.Rows.IndexOf(row)) ?? row;
        }

        // The column to put the cursor on: `wanted`, while the part shown has it.
        private int ColumnShownNear(BoardTableSheet sheet, int wanted)
        {
            int column = Math.Clamp(wanted, 0, sheet.Columns.Count - 1);
            List<int> shown = BoardTableSheetParts.ColumnsShown(this.thisPart, sheet);

            return shown.Contains(column) || shown.Count == 0 ? column : shown[0];
        }

        private void NoteSheetShown()
        {
            if (this.thisCurrentSheet is not { } sheet || this.thisSheetToldShown == (sheet.Name, this.thisPart))
            {
                return;
            }

            this.thisSheetToldShown = (sheet.Name, this.thisPart);
            this.SheetShown?.Invoke(this, EventArgs.Empty);
        }

        private void ForgetSheetShown()
        {
            this.thisSheetToldShown = null;
            this.thisPart = BoardTableSheetPart.Whole;
        }
    }
}
