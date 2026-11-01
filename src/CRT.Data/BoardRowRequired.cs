using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A ROW OF THE BOARD'S OWN LISTS MUST HAVE FILLED IN (owner requests, 2026-10-30: "'What' and
    // 'Name or handle' must be set ... In 'Links' then all 3 fields must be filled ... Same for 'Files'
    // ... the 'Important signal name' must not be empty"; and 2026-11-01: "These errors must be flagged
    // in 'Board status' and then it should not be possible to submit the draft board"):
    //
    //   - a BOARD FILE: its category, its name and its file;
    //   - a BOARD LINK: its category, its name and its address;
    //   - a CREDIT: what it is for (its category) and the name or handle - which part of it and the
    //     contact may stay empty, as on many of the live boards' credits;
    //   - an IMPORTANT SIGNAL: its KiCad net and its name.
    //
    // *** ONE RULE, ASKED EVERYWHERE. *** BoardDataChecks makes each empty one an ERROR on its cell -
    // the table marks it, the draft's row counts it, Board status says it, Submit sends nothing while it
    // is there, and the server refuses it (an error in the table is a refusal at submit); the component
    // pane marks the box and puts the row into the draft as it is (BoardPartRequired names the fields
    // by its headings). Until 2026-11-01 the pane held such a row back, so nothing but the pane knew.
    //
    // Columns are named as the board's sheets name them (BoardWorkbookSchema), in the order a row shows
    // them. A blank value is an empty one.
    // ###########################################################################################
    public static class BoardRowRequired
    {
        public static IReadOnlyList<string> EmptyColumns(BoardLocalFileEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return BoardRowRequired.Empty(
                (row.Category, BoardWorkbookSchema.ColCategory),
                (row.Name, BoardWorkbookSchema.ColName),
                (row.File, BoardWorkbookSchema.ColFile));
        }

        public static IReadOnlyList<string> EmptyColumns(BoardLinkEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return BoardRowRequired.Empty(
                (row.Category, BoardWorkbookSchema.ColCategory),
                (row.Name, BoardWorkbookSchema.ColName),
                (row.Url, BoardWorkbookSchema.ColUrl));
        }

        public static IReadOnlyList<string> EmptyColumns(CreditEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return BoardRowRequired.Empty(
                (row.Category, BoardWorkbookSchema.ColCategory),
                (row.NameOrHandle, BoardWorkbookSchema.ColNameOrHandle));
        }

        public static IReadOnlyList<string> EmptyColumns(KiCadImportantSignalEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return BoardRowRequired.Empty(
                (row.KiCadNetName, BoardWorkbookSchema.ColKiCadNetName),
                (row.DisplayName, BoardWorkbookSchema.ColDisplayName));
        }

        // ###########################################################################################
        // How many rows of each list have a column they need empty, and how many such columns there
        // are in all - each one ERROR of BoardDataChecks (a board file's empty file is its
        // "file.unreferenced", the rest "row.required_empty"), so Board status can tell them apart from
        // the board's other errors.
        // ###########################################################################################
        public static BoardRowsIncomplete Of(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            static (int Rows, int Columns) Count<T>(IEnumerable<T>? rows, Func<T, IReadOnlyList<string>> emptyColumns)
            {
                List<int> empty = (rows ?? []).Select(row => emptyColumns(row).Count).Where(count => count > 0).ToList();
                return (empty.Count, empty.Sum());
            }

            (int files, int fileColumns) = Count(board.BoardLocalFiles, BoardRowRequired.EmptyColumns);
            (int links, int linkColumns) = Count(board.BoardLinks, BoardRowRequired.EmptyColumns);
            (int credits, int creditColumns) = Count(board.Credits, BoardRowRequired.EmptyColumns);
            (int signals, int signalColumns) = Count(board.KiCadImportantSignals, BoardRowRequired.EmptyColumns);

            return new BoardRowsIncomplete(files, links, credits, signals, fileColumns + linkColumns + creditColumns + signalColumns);
        }

        private static List<string> Empty(params (string? Value, string Column)[] columns) =>
            columns.Where(column => string.IsNullOrWhiteSpace(column.Value)).Select(column => column.Column).ToList();
    }

    // The rows of each of the board's own lists with a column they need empty, and the errors they make.
    public sealed record BoardRowsIncomplete(int Files, int Links, int Credits, int Signals, int Errors)
    {
        public bool Any => this.Errors > 0;
    }
}
