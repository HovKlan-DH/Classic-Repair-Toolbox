using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A SHEET SHOWN IN TWO (owner request, 2026-11-03: "the table view of 'Component images' should
    // probably be divided into two now, as it now has 'Pictures' and 'Oscilloscope baseline'
    // seperately"). The sheet stays ONE sheet, as the board file and the server keep it: this only says
    // which of its rows and columns the table shows for each of the two entries of the Contribute tab's left
    // column, and each of the two tabs of the Maintainer tab's table (BoardTableTabs). A row is a baseline by its pin alone - ComponentImageKinds' rule, read here off the row's cell, so a pin
    // typed or cleared moves the row across.
    //
    // A row with nothing typed in it is shown in either, as under the pills: it is somebody's new row,
    // and a filter hiding it the moment it appeared would read as the insert having failed.
    //
    // A picture has no pin and none of what goes with one, so those columns are left out of its part.
    // Whatever walks along the columns - Tab, a pasted block - goes by ColumnsShown, or it writes into
    // a cell nobody can see (a pin, which would make the picture a baseline). What the pills and the search
    // judge there is CellsShown (code review, 2026-11-04): a row listed for a word or a problem in a column
    // it hides shows nothing that says why, and the cursor sent to such a column went nowhere.
    //
    // *** A ROW THAT CAN STILL BE A BASELINE ALSO SHOWS AMONG THE BASELINES, WHERE ITS PIN CAN BE TYPED (code
    // review, 2026-11-04). *** Neither "Pictures" here nor the pane's page has a pin field, so a row that
    // was only a picture by then was out of reach: "Insert row", a label typed, another row clicked, and only
    // Ctrl+Z or deleting it got out. That is a row worked on among the baselines, while this table stands
    // (BoardTableRow.StaysWithBaselines), one holding what only a baseline has, or one with nothing yet that
    // says picture. It is still a picture by its pin alone, as for the rest of CRT, and shows among the
    // pictures too, so that table loses nothing. Measured on the published data, 2026-11-04: 15 of the 2,784
    // rows with no pin hold such a value - nine "Failure Mode" rows of the CPC 664, six "Pinout" rows of the
    // VIC-20's two boards - now in sight there; none of the 2,784 has neither a file nor a note.
    // ###########################################################################################
    public enum BoardTableSheetPart
    {
        Whole,
        Pictures,
        Baselines,
    }

    public static class BoardTableSheetParts
    {
        private static readonly HashSet<string> BaselineOnlyColumns = new(StringComparer.Ordinal)
        {
            BoardWorkbookSchema.ColPin,
            BoardWorkbookSchema.ColExpectedOscilloscopeReading,
            BoardWorkbookSchema.ColTimeDiv,
            BoardWorkbookSchema.ColVoltsDiv,
            BoardWorkbookSchema.ColTriggerLevelVolts,
            BoardWorkbookSchema.ColMachineState,
        };

        public static bool IsSplit(string? sheetName) =>
            string.Equals(sheetName, BoardWorkbookSchema.SheetComponentImages, StringComparison.Ordinal);

        // Which part a row of a sheet in two belongs to.
        public static BoardTableSheetPart PartOf(BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return row.Sheet.CellText(row, BoardWorkbookSchema.ColPin).Length > 0
                ? BoardTableSheetPart.Baselines
                : BoardTableSheetPart.Pictures;
        }

        public static bool Shows(BoardTableSheetPart part, BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return part == BoardTableSheetPart.Whole ||
                   row.State is BoardTableRowState.Blank ||
                   BoardTableSheetParts.PartOf(row) == part ||
                   (part == BoardTableSheetPart.Baselines && BoardTableSheetParts.MayStillBeABaseline(row));
        }

        // A row with no pin that is to be shown where it can be given one - see the header.
        private static bool MayStillBeABaseline(BoardTableRow row)
        {
            if (row.StaysWithBaselines)
            {
                return true;
            }

            bool saysPicture = false;

            foreach (BoardTableCell cell in row.Cells)
            {
                if (cell.Text.Length > 0)
                {
                    string column = row.Sheet.Columns[cell.ColumnIndex];

                    if (BoardTableSheetParts.BaselineOnlyColumns.Contains(column))
                    {
                        return true;
                    }

                    saysPicture |= column is BoardWorkbookSchema.ColFile or BoardWorkbookSchema.ColNote;
                }
            }

            return !saysPicture;
        }

        // The cells of `row` the part shows - what its pills and search go by.
        public static IEnumerable<BoardTableCell> CellsShown(BoardTableSheetPart part, BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return part == BoardTableSheetPart.Pictures
                ? row.Cells.Where(cell => !BoardTableSheetParts.BaselineOnlyColumns.Contains(row.Sheet.Columns[cell.ColumnIndex]))
                : row.Cells;
        }

        // The columns of `sheet` the part shows, by their index in the sheet - every one but a picture's.
        public static List<int> ColumnsShown(BoardTableSheetPart part, BoardTableSheet sheet)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            return Enumerable.Range(0, sheet.Columns.Count)
                .Where(index => part != BoardTableSheetPart.Pictures || !BoardTableSheetParts.BaselineOnlyColumns.Contains(sheet.Columns[index]))
                .ToList();
        }

        // ###########################################################################################
        // The part of `sheet` to show: `wanted` - while it shows something of what the pills and the
        // search ask for, or the other part shows nothing either. So picking "Errors" goes to the errors
        // inside a sheet in two as it does between sheets. Any other sheet is whole.
        // ###########################################################################################
        public static BoardTableSheetPart PartToShow(BoardTableSheet sheet, BoardTableSheetPart wanted, BoardTableRowKinds kinds, BoardTableSearch search)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(search);

            if (!BoardTableSheetParts.IsSplit(sheet.Name))
            {
                return BoardTableSheetPart.Whole;
            }

            BoardTableSheetPart part = wanted == BoardTableSheetPart.Whole ? BoardTableSheetPart.Pictures : wanted;

            if (kinds == BoardTableRowKinds.None && !search.IsActive)
            {
                return part;
            }

            BoardTableSheetPart other = part == BoardTableSheetPart.Pictures ? BoardTableSheetPart.Baselines : BoardTableSheetPart.Pictures;

            return !BoardTableSheetParts.ShowsAnything(sheet, part, kinds, search) && BoardTableSheetParts.ShowsAnything(sheet, other, kinds, search)
                ? other
                : part;
        }

        // ###########################################################################################
        // How many of a sheet's changes a part holds - its tab's count (owner request, 2026-11-05). By the
        // row's state, the sheet's own count being the sum: a row added, changed or deleted, each in the part
        // its pin puts it. A sheet not in two is its own count.
        // ###########################################################################################
        public static int ChangeCountOf(BoardTableSheet sheet, BoardTableSheetPart part)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            return part == BoardTableSheetPart.Whole
                ? sheet.ChangeCount
                : sheet.Rows.Count(row => row.State is BoardTableRowState.Added or BoardTableRowState.Modified or BoardTableRowState.Deleted &&
                                          BoardTableSheetParts.PartOf(row) == part);
        }

        // Whether the pills and the search leave anything of this part of the sheet to see.
        public static bool ShowsAnything(BoardTableSheet sheet, BoardTableSheetPart part, BoardTableRowKinds kinds, BoardTableSearch search)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(search);

            return sheet.Rows.Any(row => row.State is not BoardTableRowState.Blank &&
                                         BoardTableSheetParts.Shows(part, row) &&
                                         BoardTableRowFilter.Shows(kinds, search, rowInHand: null, row, part));
        }

        // Whether `sheet` holds a row with the values `row` has, in the same state - a row a step left as
        // it stood, wherever that step moved it to.
        public static bool StandsIn(BoardTableSheet sheet, BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(row);

            return sheet.Rows.Any(candidate => candidate.IsDeleted == row.IsDeleted &&
                                               candidate.Cells.Select(cell => cell.Text).SequenceEqual(row.Cells.Select(cell => cell.Text)));
        }

        // The row at `index` or nearest it the part shows - from there down, then up. Null with none.
        public static BoardTableRow? NearestShown(BoardTableSheet sheet, BoardTableSheetPart part, int index)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            int from = Math.Clamp(index, 0, Math.Max(sheet.Rows.Count - 1, 0));

            return sheet.Rows.Skip(from).FirstOrDefault(row => BoardTableSheetParts.Shows(part, row))
                ?? sheet.Rows.Take(from).LastOrDefault(row => BoardTableSheetParts.Shows(part, row));
        }
    }
}
