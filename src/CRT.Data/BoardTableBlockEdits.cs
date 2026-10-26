using System;
using System.Collections.Generic;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A PASTE, A FILL DOWN OR A REPLACE DID, AND HOW THE TABLE SAYS SO (CRT 4.0, 2026-10-25 -
    // the plan's step 0: "the table view gains block copy and paste, fill down, and find &
    // replace"). The rules are BoardTableSheet.PasteBlock / PasteIntoCells / FillDown and
    // BoardTableDocument.ReplaceAll; the sentences are here, in one place, so the Drafts tab and the
    // Maintainer tab say the same and the words are tested.
    //
    // Every sentence for a change ends by saying Ctrl+Z undoes it: each of these writes many cells
    // at once, and the one thing to know after a surprise is that it is one step back.
    // ###########################################################################################

    // What PasteBlock did. `Rows` are the rows written into, in order (the ones added among them),
    // `FirstColumn` and `ColumnCount` the columns - what the editor selects afterwards.
    public sealed record BoardTablePasteResult(
        IReadOnlyList<BoardTableRow> Rows,
        int FirstColumn,
        int ColumnCount,
        int CellsChanged,
        int RowsAdded,
        int RowsLeftOut,
        int ColumnsLeftOut,
        bool HeadingsLeftOut)
    {
        public static BoardTablePasteResult Nothing { get; } = new([], 0, 0, 0, 0, 0, 0, false);

        public bool Pasted => this.Rows.Count > 0;
    }

    // What ReplaceAll did: how many times the text was replaced, in how many cells, on how many sheets.
    public sealed record BoardTableReplaceResult(int Occurrences, int Cells, int Sheets);

    public static class BoardTableEditWording
    {
        public const string UndoHint = "Ctrl+Z undoes it.";

        // ###########################################################################################
        // A paste of several cells: where it went, then whatever of the clipboard did not go in and
        // why. Empty for ONE cell pasted into one cell, which needs no words - the cell shows it.
        // ###########################################################################################
        public static string Pasted(BoardTablePasteResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            if (!result.Pasted ||
                (result.Rows.Count == 1 && result.ColumnCount == 1 && !BoardTableEditWording.AnythingLeftOut(result)))
            {
                return string.Empty;
            }

            var parts = new List<string>();

            if (result.CellsChanged == 0 && result.RowsAdded == 0)
            {
                parts.Add("The cells already held what was pasted - nothing changed.");
            }
            else
            {
                parts.Add($"Pasted into {Count(result.Rows.Count, "row")}.");

                if (result.RowsAdded > 0)
                {
                    parts.Add(result.RowsAdded == 1
                        ? "1 of them was added at the end of the sheet."
                        : $"{Number(result.RowsAdded)} of them were added at the end of the sheet.");
                }
            }

            if (result.HeadingsLeftOut)
            {
                parts.Add("The first row copied was this sheet's column headings, so it was left out.");
            }

            if (result.ColumnsLeftOut > 0)
            {
                parts.Add(result.ColumnsLeftOut == 1
                    ? "1 column did not fit and was left out."
                    : $"{Number(result.ColumnsLeftOut)} columns did not fit and were left out.");
            }

            if (result.RowsLeftOut > 0)
            {
                parts.Add(
                    $"{(result.RowsLeftOut == 1 ? "1 row was" : $"{Number(result.RowsLeftOut)} rows were")} left out: while the " +
                    "table shows only some rows, a paste goes into those and adds none.");
            }

            if (result.CellsChanged > 0 || result.RowsAdded > 0)
            {
                parts.Add(BoardTableEditWording.UndoHint);
            }

            return string.Join(" ", parts);
        }

        // One value pasted onto several selected cells.
        public static string PastedIntoCells(int changed) => changed == 0
            ? "The cells already held what was pasted - nothing changed."
            : $"Pasted into {Count(changed, "cell")}. {BoardTableEditWording.UndoHint}";

        public const string NothingToPaste = "The clipboard holds no text to paste.";

        // ###########################################################################################
        // Nothing can be written into a red row. Said without "your draft" (code review, 2026-10-25):
        // the Maintainer tab shows the same table, a submission's and BETA's.
        // ###########################################################################################
        public const string PasteOnDeletedRow = "A red row has been deleted, so nothing can be pasted into it. Ctrl+Z brings it back if you deleted it.";

        public const string PasteOnDeletedRows = "The selected cells are all on red rows, which have been deleted, so nothing was pasted. Ctrl+Z brings a row back if you deleted it.";

        public const string FillOnDeletedRows = "The cells to fill are all on red rows, which have been deleted, so nothing was filled. Ctrl+Z brings a row back if you deleted it.";

        public static string Filled(int changed) => changed == 0
            ? "Nothing to fill - the cells already hold the value above them."
            : $"Filled {Count(changed, "cell")} with the value at the top. {BoardTableEditWording.UndoHint}";

        public const string NothingAbove = "There is no row above to fill from. Select several rows to fill the top one's values down.";

        // ###########################################################################################
        // A replace: how often, or that nothing on the sheets searched contains the text. `everySheet`
        // says which sheets were searched - the one on screen, or all of them.
        // ###########################################################################################
        public static string Replaced(BoardTableReplaceResult result, string find, bool everySheet)
        {
            ArgumentNullException.ThrowIfNull(result);

            if (result.Cells == 0)
            {
                return everySheet
                    ? $"No row shown on any sheet contains [{find}]."
                    : $"No row shown on this sheet contains [{find}].";
            }

            string where = everySheet && result.Sheets > 1 ? $" on {Number(result.Sheets)} sheets" : string.Empty;

            return $"Replaced [{find}] {Count(result.Occurrences, "time")} in {Count(result.Cells, "cell")}{where}. {BoardTableEditWording.UndoHint}";
        }

        public const string NothingToFind = "Type the text to find first.";

        private static bool AnythingLeftOut(BoardTablePasteResult result) =>
            result.HeadingsLeftOut || result.ColumnsLeftOut > 0 || result.RowsLeftOut > 0 || result.RowsAdded > 0;

        private static string Count(int count, string noun) =>
            count == 1 ? $"1 {noun}" : $"{Number(count)} {noun}s";

        private static string Number(int count) => count.ToString("N0", CultureInfo.InvariantCulture);
    }
}
