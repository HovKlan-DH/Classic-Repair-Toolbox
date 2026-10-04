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
    // *** A NEW ROW STILL BEING FILLED IN IS ALWAYS SHOWN. *** It is the contributor's own work in
    // progress - just inserted, not typed in yet, or not yet complete enough to be saved - and a
    // filter hiding it the moment it appeared would read as the insert having failed. "Show changes
    // only" showed it for the same reason.
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

        // The kinds one row is - several at once: a changed row can carry an error too.
        public static BoardTableRowKinds KindsOf(BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            BoardTableRowKinds kinds = row.State switch
            {
                BoardTableRowState.Added => BoardTableRowKinds.Added,
                BoardTableRowState.Modified => BoardTableRowKinds.Modified,
                BoardTableRowState.Deleted => BoardTableRowKinds.Deleted,
                _ => BoardTableRowKinds.None
            };

            if (row.HasErrors)
                kinds |= BoardTableRowKinds.Errors;

            if (row.HasWarnings)
                kinds |= BoardTableRowKinds.Warnings;

            return kinds;
        }

        // Whether `kinds` shows `row` - every row when nothing is picked, and always a new one still
        // being filled in.
        public static bool Shows(BoardTableRowKinds kinds, BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (kinds == BoardTableRowKinds.None || row.State is BoardTableRowState.Blank or BoardTableRowState.Incomplete)
                return true;

            return (BoardTableRowFilter.KindsOf(row) & kinds) != BoardTableRowKinds.None;
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
