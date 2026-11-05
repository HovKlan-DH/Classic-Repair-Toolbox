using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH ROWS THE TABLE SHOWS (owner request, 2026-10-02: the colour key's counts are clickable,
    // "and even do remove the "Show changes only" so it works in the same unified way").
    //
    // Each count in the colour key is a kind of row - Added, Modified, Deleted, and since the same
    // day Errors and Warnings (BoardDataChecks). Clicking counts picks kinds; the table then shows
    // only rows of ANY picked kind, and hides the tabs of sheets that have none
    // (BoardTableDocument.SheetsShown). Nothing picked shows every row. "Show changes only" is what
    // an old remembered choice becomes (Changes).
    //
    // *** THERE WAS A SIXTH, FLAGGED (a duplicate row, or one the save leaves out). *** Both are
    // WARNINGS since 2026-10-03 (owner request: "Should flagged now be treated as warnings?"), so
    // a choice remembered with "Flagged" in it reads as Warnings (Parse), and its number, 8, is
    // left unused rather than given to anything else.
    //
    // *** A NEW ROW WITH NOTHING TYPED IN IT IS ALWAYS SHOWN. *** It is the contributor's own work in
    // progress, and a filter hiding it the moment it appeared would read as the insert having failed.
    // "Show changes only" showed it for the same reason.
    //
    // *** AND SO IS THE ROW THE CURSOR IS ON, UNTIL THE CURSOR LEAVES IT (owner's answer, 2026-11-02).
    // *** What the person is working in may stop being of a picked kind with the very text just typed:
    // a new row no longer blank under the Modified pill, a value typed back to the published one, a
    // row's last error fixed under the Errors pill. Taken away there and then, the cursor had to land
    // on another row, and the next text typed overwrote that row's cell. A row with no cursor in it
    // still goes at once. A pick or a search, being a new question, keeps nothing (the editor's
    // thisRowInHand, BoardTableEditor.Filter.cs).
    // ###########################################################################################
    [Flags]
    public enum BoardTableRowKinds
    {
        None = 0,
        Added = 1,
        Modified = 2,
        Deleted = 4,
        Errors = 16,
        Warnings = 32
    }

    public static class BoardTableRowFilter
    {
        // What "Show changes only" showed, before the counts became the filter - less Flagged, which
        // is warnings now.
        public const BoardTableRowKinds Changes =
            BoardTableRowKinds.Added | BoardTableRowKinds.Modified | BoardTableRowKinds.Deleted;

        // The name a remembered choice used for what is Warnings now.
        private const string FlaggedName = "Flagged";

        // The kinds one row is - several at once: a changed row can carry an error too. On a part of a sheet
        // in two, only a problem in a column it shows counts (BoardTableSheetParts.CellsShown).
        public static BoardTableRowKinds KindsOf(BoardTableRow row, BoardTableSheetPart part = BoardTableSheetPart.Whole)
        {
            ArgumentNullException.ThrowIfNull(row);

            BoardTableRowKinds kinds = row.State switch
            {
                BoardTableRowState.Added => BoardTableRowKinds.Added,
                BoardTableRowState.Modified => BoardTableRowKinds.Modified,
                BoardTableRowState.Deleted => BoardTableRowKinds.Deleted,
                _ => BoardTableRowKinds.None
            };

            foreach (BoardTableCell cell in BoardTableSheetParts.CellsShown(part, row))
            {
                foreach (BoardDataProblem problem in cell.Problems)
                {
                    kinds |= problem.Level switch
                    {
                        BoardProblemLevel.Error => BoardTableRowKinds.Errors,
                        BoardProblemLevel.Warning => BoardTableRowKinds.Warnings,
                        _ => BoardTableRowKinds.None,
                    };
                }
            }

            return kinds;
        }

        // Whether `kinds` shows `row` - every row when nothing is picked, and always a new one still
        // being filled in.
        public static bool Shows(BoardTableRowKinds kinds, BoardTableRow row, BoardTableSheetPart part = BoardTableSheetPart.Whole)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (kinds == BoardTableRowKinds.None || row.State is BoardTableRowState.Blank)
                return true;

            return (BoardTableRowFilter.KindsOf(row, part) & kinds) != BoardTableRowKinds.None;
        }

        // Whether the table shows `row`: the row the cursor is on (see the header), or one both the
        // picked pills and the search show - on `part`, by the columns it shows.
        public static bool Shows(BoardTableRowKinds kinds, BoardTableSearch search, BoardTableRow? rowInHand, BoardTableRow row, BoardTableSheetPart part = BoardTableSheetPart.Whole)
        {
            ArgumentNullException.ThrowIfNull(search);

            return ReferenceEquals(row, rowInHand) || (BoardTableRowFilter.Shows(kinds, row, part) && search.Shows(row, part));
        }

        // The rows of a sheet the table shows, in the sheet's order - of `part`, on a sheet shown in two
        // (BoardTableSheetParts). The row in hand stays whichever part it now belongs to.
        public static List<BoardTableRow> RowsShown(
            IEnumerable<BoardTableRow> rows,
            BoardTableRowKinds kinds,
            BoardTableSearch search,
            BoardTableRow? rowInHand,
            BoardTableSheetPart part = BoardTableSheetPart.Whole)
        {
            ArgumentNullException.ThrowIfNull(rows);

            return rows
                .Where(row => ReferenceEquals(row, rowInHand) || BoardTableSheetParts.Shows(part, row))
                .Where(row => BoardTableRowFilter.Shows(kinds, search, rowInHand, row, part))
                .ToList();
        }

        // ###########################################################################################
        // The cell of `row` to work on next when the Errors or Warnings pill is what brought it into
        // view: its first with a problem of a picked level. Null for a row with none, and with neither
        // pill picked - a warning means nothing to someone going through the rows Modified shows.
        //
        // Where the cursor starts when the table is given the focus under the pill (code review,
        // 2026-11-02): Submit on a draft with errors opens it so, and what is typed then goes where
        // something is wrong anyway. Only a column `part` shows: the grid has no other to put it on.
        // ###########################################################################################
        public static BoardTableCell? FirstProblemCell(BoardTableRowKinds kinds, BoardTableRow row, BoardTableSheetPart part = BoardTableSheetPart.Whole)
        {
            ArgumentNullException.ThrowIfNull(row);

            bool errors = kinds.HasFlag(BoardTableRowKinds.Errors);
            bool warnings = kinds.HasFlag(BoardTableRowKinds.Warnings);

            return BoardTableSheetParts.CellsShown(part, row).FirstOrDefault(cell => cell.Problems.Any(problem =>
                (errors && problem.Level == BoardProblemLevel.Error) ||
                (warnings && problem.Level == BoardProblemLevel.Warning)));
        }

        // ###########################################################################################
        // The choice as a setting, and back - kind names joined by commas, so a settings file stays
        // readable and a kind added later cannot shift what an old number meant. Unknown names are
        // skipped.
        // ###########################################################################################
        public static string Format(BoardTableRowKinds kinds) =>
            string.Join(",", Enum.GetValues<BoardTableRowKinds>()
                .Where(kind => kind != BoardTableRowKinds.None && kinds.HasFlag(kind))
                .Select(kind => kind.ToString()));

        public static BoardTableRowKinds Parse(string? text)
        {
            BoardTableRowKinds kinds = BoardTableRowKinds.None;

            foreach (string name in (text ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (string.Equals(name, BoardTableRowFilter.FlaggedName, StringComparison.OrdinalIgnoreCase))
                    kinds |= BoardTableRowKinds.Warnings;
                else if (Enum.TryParse(name, ignoreCase: true, out BoardTableRowKinds kind) && Enum.IsDefined(kind))
                    kinds |= kind;
            }

            return kinds;
        }

        // The kinds, one at a time, in the colour key's order.
        public static IReadOnlyList<BoardTableRowKinds> Each { get; } =
        [
            BoardTableRowKinds.Added,
            BoardTableRowKinds.Modified,
            BoardTableRowKinds.Deleted,
            BoardTableRowKinds.Errors,
            BoardTableRowKinds.Warnings
        ];
    }
}
