using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE TABLE'S FILE CELLS - which cells name a file, and which file each side of the comparison
    // is (owner request, 2026-09-26: "whenever it is a file column, and it is an image, then it
    // should show a hover-helper with the image ... both the removed and the new image side-by-side
    // ... If the file is something else, e.g. PDF, then there should be a link that will open").
    //
    // In CRT.Data, like every other rule of the table, so the Drafts tab and the maintainer
    // application show the same thing: BoardTableEditor only paints the answer, and each host only
    // fetches the bytes (IBoardTableFileSource).
    // ###########################################################################################
    public static class BoardTableFileCells
    {
        // ###########################################################################################
        // Whether a column holds a data-root-relative file path. Four columns do: the schematic
        // image of "Board schematics", and "File" on the three sheets that attach files.
        // ###########################################################################################
        public static bool IsFileColumn(string? sheetName, string? columnName)
        {
            if (string.Equals(sheetName, BoardWorkbookSchema.SheetBoardSchematics, StringComparison.Ordinal))
            {
                return string.Equals(columnName, BoardWorkbookSchema.ColSchematicImageFile, StringComparison.Ordinal);
            }

            bool isFileSheet =
                string.Equals(sheetName, BoardWorkbookSchema.SheetComponentImages, StringComparison.Ordinal) ||
                string.Equals(sheetName, BoardWorkbookSchema.SheetComponentLocalFiles, StringComparison.Ordinal) ||
                string.Equals(sheetName, BoardWorkbookSchema.SheetBoardLocalFiles, StringComparison.Ordinal);

            return isFileSheet && string.Equals(columnName, BoardWorkbookSchema.ColFile, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The file each side of a file cell names, or null for a cell that names none (not a file
        // column, or blank on both sides).
        //
        //   - a live row paired with a published row: the published row's value, and this cell's;
        //   - an added row (or one with nothing to pair with): no published side;
        //   - a deleted row's ghost: the published side only - it shows what was published.
        //
        // The same path on both sides is still both sides: the file at that path can have been
        // REPLACED (same name, new picture), which only the bytes can tell - see IsSamePath.
        // ###########################################################################################
        public static BoardTableFileCell? Of(BoardTableCell? cell)
        {
            if (cell is null)
            {
                return null;
            }

            BoardTableRow row = cell.Row;

            if (!BoardTableFileCells.IsFileColumn(row.Sheet.Name, row.Sheet.Columns[cell.ColumnIndex]))
            {
                return null;
            }

            string? current = row.IsDeleted ? null : BoardTableFileCells.NonBlank(cell.Text);

            string? published = row.IsDeleted
                ? BoardTableFileCells.NonBlank(cell.Text)
                : row.PublishedIndex >= 0 ? BoardTableFileCells.NonBlank(cell.PublishedText) : null;

            return published is null && current is null
                ? null
                : new BoardTableFileCell(published, current);
        }

        // ###########################################################################################
        // The line above the preview's pictures, or null for none. `sameContent` is whether the two
        // sides' bytes are identical - known only once both are read, and only asked about when both
        // sides name the same path.
        // ###########################################################################################
        public static string? Headline(BoardTableFileCell file, bool? sameContent)
        {
            ArgumentNullException.ThrowIfNull(file);

            if (file.PublishedPath is null)
                return "New file";

            if (file.CurrentPath is null)
                return "Removed";

            if (!file.IsSamePath)
                return "Changed to another file";

            return sameContent switch
            {
                true => "Unchanged",
                false => "Replaced - same file name, new content",
                _ => null
            };
        }

        private static string? NonBlank(string? value)
        {
            string trimmed = value?.Trim() ?? string.Empty;
            return trimmed.Length == 0 ? null : trimmed;
        }
    }

    // ###########################################################################################
    // One file cell's two sides: the path as published (null when there is none) and the path as
    // this version names it (null for a deleted row, or a cell emptied).
    // ###########################################################################################
    public sealed record BoardTableFileCell(string? PublishedPath, string? CurrentPath)
    {
        // Both sides name the same path - so whether it CHANGED is a question for the bytes.
        // Ordinal: the server's tree is case-sensitive, so "U8.png" and "u8.png" are two files.
        public bool IsSamePath =>
            this.PublishedPath is not null &&
            string.Equals(this.PublishedPath, this.CurrentPath, StringComparison.Ordinal);
    }
}
