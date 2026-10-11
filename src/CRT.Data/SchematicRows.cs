using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One row of a board's "Board schematics" sheet, as a window listing the sheet points at it: its
    // name, its image file, and - among rows sharing both - which of them, counted from the top.
    // ###########################################################################################
    public sealed record SchematicRowRef(string SchematicName, string ImageFile, int Occurrence = 0);

    // ###########################################################################################
    // FINDING THE ROW A CHANGE IS FOR, BY MORE THAN ITS NAME (code review, 2026-10-10).
    //
    // The "Schematic images" window saved every change - a rename, a CAD name, a removal - to the
    // FIRST row of the image's name. The name is the sheet's key, but two rows CAN share one (typed
    // in Excel, an error the checks report), and renaming one of them is exactly how a contributor
    // fixes that: the change landed on the other image. A removal took both.
    //
    // So a row is found by its name AND its image file, and where even those are shared, by its
    // place among the rows sharing them. The draft is read again for every change; a row it no longer
    // has (changed in Excel meanwhile) is not found, and the window reads the list again.
    // ###########################################################################################
    public static class SchematicRows
    {
        // The reference for the row at `index` of a list shown in the sheet's order.
        public static SchematicRowRef RefOf(IReadOnlyList<(string SchematicName, string ImageFile)> rows, int index)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, rows.Count);

            (string name, string file) = rows[index];

            int occurrence = rows
                .Take(index)
                .Count(row => SchematicRename.SameName(row.SchematicName, name) && SchematicRows.SameImageFile(row.ImageFile, file));

            return new SchematicRowRef(name ?? string.Empty, file ?? string.Empty, occurrence);
        }

        // Where the row is in `schematics`, or -1 when it is not there.
        public static int IndexOf(IReadOnlyList<BoardSchematicEntry> schematics, SchematicRowRef row)
        {
            int seen = 0;

            for (int i = 0; i < schematics.Count; i++)
            {
                if (!SchematicRename.SameName(schematics[i].SchematicName, row.SchematicName) ||
                    !SchematicRows.SameImageFile(schematics[i].SchematicImageFile, row.ImageFile))
                {
                    continue;
                }

                if (seen == row.Occurrence)
                {
                    return i;
                }

                seen++;
            }

            return -1;
        }

        // ###########################################################################################
        // The board with the row's CAD name replaced - every other column of it kept
        // (BoardSchematicEntry.WithCadName), and every other row - and whether the row was there.
        // ###########################################################################################
        public static (BoardData Board, bool Found) WithCadName(BoardData board, SchematicRowRef row, string cadName)
        {
            int index = SchematicRows.IndexOf(board.Schematics, row);

            if (index < 0)
            {
                return (board, false);
            }

            var schematics = new List<BoardSchematicEntry>(board.Schematics);
            schematics[index] = schematics[index].WithCadName(cadName?.Trim() ?? string.Empty);

            return (board.WithSchematics(schematics), true);
        }

        // The board without the row - only that row, never another of its name - and whether it was there.
        public static (BoardData Board, bool Found) Without(BoardData board, SchematicRowRef row)
        {
            int index = SchematicRows.IndexOf(board.Schematics, row);

            if (index < 0)
            {
                return (board, false);
            }

            var schematics = new List<BoardSchematicEntry>(board.Schematics);
            schematics.RemoveAt(index);

            return (board.WithSchematics(schematics), true);
        }

        // Two image paths name one file: trimmed, either slash, without regard to capitals. Unlike
        // SchematicRenameLog.SameFile, two blank paths are the same here - two rows naming no file
        // are told apart by their place among the rows sharing their name.
        public static bool SameImageFile(string? first, string? second) =>
            string.Equals(
                (first ?? string.Empty).Trim().Replace('\\', '/'),
                (second ?? string.Empty).Trim().Replace('\\', '/'),
                StringComparison.OrdinalIgnoreCase);
    }
}
