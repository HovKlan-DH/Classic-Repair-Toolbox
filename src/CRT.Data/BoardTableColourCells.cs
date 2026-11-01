using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE TABLE'S COLOR CELLS - which cells hold a color, and what a changed one changed from and to
    // (owner request, 2026-11-01: "If the color has changed, then I would like a popup showing the colors
    // 'Changed from' and 'Changed to'").
    //
    // In CRT.Data, like every other rule of the table, so the Contribute and Maintainer tabs show the
    // same thing: BoardTableEditor only paints the answer.
    // ###########################################################################################
    public static class BoardTableColourCells
    {
        // The four color columns of "Board schematics" - an opacity is a number, and says itself.
        public static bool IsColourColumn(string? sheetName, string? columnName) =>
            string.Equals(sheetName, BoardWorkbookSchema.SheetBoardSchematics, StringComparison.Ordinal) &&
            columnName is BoardWorkbookSchema.ColSchematicHighlightColor
                or BoardWorkbookSchema.ColOppositeTraceHighlightColor
                or BoardWorkbookSchema.ColThumbnailHighlightColor
                or BoardWorkbookSchema.ColTraceHighlightColor;

        // The two sides of a color cell that changed - null for any other cell. An added row and a deleted
        // row's ghost have no such cell: only a cell paired with a published row is ever Modified.
        public static BoardTableColourChange? Of(BoardTableCell? cell)
        {
            if (cell is null || cell.State != BoardTableCellState.Modified)
            {
                return null;
            }

            BoardTableRow row = cell.Row;

            return BoardTableColourCells.IsColourColumn(row.Sheet.Name, row.Sheet.Columns[cell.ColumnIndex])
                ? new BoardTableColourChange(cell.PublishedText.Trim(), cell.Text.Trim())
                : null;
        }
    }

    // One changed color cell's two sides, as the board keeps them - blank when a side has no color of its own.
    public sealed record BoardTableColourChange(string From, string To);
}
