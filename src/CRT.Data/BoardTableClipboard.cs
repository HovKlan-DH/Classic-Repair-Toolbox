using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // CELLS TO AND FROM THE CLIPBOARD, the way Excel, LibreOffice and Google Sheets read and write
    // them - for the table editor's Ctrl+C / Ctrl+V (owner request, 2026-09-24).
    //
    // *** ONE CELL AT A TIME UNTIL CRT 4.0 (2026-10-25). *** The project owner asked for the simple
    // version first, and a clipboard holding a BLOCK of cells was refused. With the board data out
    // of Excel, the table is where every edit is made, so a block is read (ReadBlock) and written
    // (ForCells) too - "without these, dropping Excel would be a step backwards" (the 4.0 plan,
    // step 0). Where a block lands is BoardTableSheet.PasteBlock's rule.
    //
    // A spreadsheet's clipboard text, which is what these rules follow:
    //   - every copied cell, even a single one, is followed by a line break;
    //   - cells are separated by tabs and rows by line breaks;
    //   - a cell whose text contains a tab, a line break or a double quote is wrapped in double
    //     quotes, with any quote inside doubled ("" for ").
    // ###########################################################################################
    public static class BoardTableClipboard
    {
        // ###########################################################################################
        // The clipboard's cells, row by row - one cell for a single cell (TryReadSingleCell's rule),
        // nothing for an empty clipboard. A row may hold fewer cells than another; a paste then
        // writes only the cells it has.
        //
        // A cell starting with a quote is a QUOTED cell only when its closing quote is followed by
        // a tab, a line break or the end - the way a spreadsheet writes one. Anything else
        // ("5.25" drive, typed in a text editor) is taken as it is, quotes and all.
        // ###########################################################################################
        public static IReadOnlyList<IReadOnlyList<string>> ReadBlock(string? clipboardText)
        {
            if (BoardTableClipboard.TryReadSingleCell(clipboardText, out string single))
            {
                return [[single]];
            }

            if (string.IsNullOrEmpty(clipboardText))
            {
                return [];
            }

            string text = BoardTableClipboard.WithoutOneTrailingLineBreak(clipboardText);
            var rows = new List<IReadOnlyList<string>>();
            var row = new List<string>();
            int at = 0;

            while (true)
            {
                string cell;

                if (at < text.Length && text[at] == '"' && BoardTableClipboard.TryReadQuoted(text, at, out string quoted, out int next))
                {
                    cell = quoted;
                    at = next;
                }
                else
                {
                    int end = text.IndexOfAny(['\t', '\r', '\n'], at);
                    end = end < 0 ? text.Length : end;
                    cell = text[at..end];
                    at = end;
                }

                row.Add(cell);

                if (at >= text.Length)
                {
                    rows.Add(row);
                    return rows;
                }

                if (text[at] == '\t')
                {
                    at++;
                    continue;
                }

                // A line break - "\r\n", "\n" or "\r" - ends the row.
                at += text[at] == '\r' && at + 1 < text.Length && text[at + 1] == '\n' ? 2 : 1;
                rows.Add(row);
                row = new List<string>();
            }
        }

        // ###########################################################################################
        // The clipboard text for the selected cells: one cell as ForSingleCell writes it; several as
        // a block, in the order `rowOrder` shows the rows and the columns' own order. Only the rows
        // and the columns holding a selected cell are written, and a cell of them that is not
        // selected is written empty - so a rectangle (the usual selection) comes out as it is, and
        // cells picked one by one with Ctrl+click keep their places without the rows between them.
        // ###########################################################################################
        public static string ForCells(IReadOnlyList<BoardTableCell> cells, IReadOnlyList<BoardTableRow> rowOrder)
        {
            ArgumentNullException.ThrowIfNull(cells);
            ArgumentNullException.ThrowIfNull(rowOrder);

            if (cells.Count == 0)
            {
                return string.Empty;
            }

            if (cells.Count == 1)
            {
                return BoardTableClipboard.ForSingleCell(cells[0].Text);
            }

            var position = new Dictionary<BoardTableRow, int>(ReferenceEqualityComparer.Instance);
            for (int i = 0; i < rowOrder.Count; i++)
            {
                position.TryAdd(rowOrder[i], i);
            }

            List<BoardTableRow> rows = cells
                .Select(cell => cell.Row)
                .Distinct(ReferenceEqualityComparer.Instance)
                .Cast<BoardTableRow>()
                .OrderBy(row => position.TryGetValue(row, out int at) ? at : int.MaxValue)
                .ToList();

            List<int> columns = cells.Select(cell => cell.ColumnIndex).Distinct().Order().ToList();
            var selected = new HashSet<BoardTableCell>(cells, ReferenceEqualityComparer.Instance);

            return BoardTableClipboard.ForBlock(rows
                .Select(row => (IReadOnlyList<string>)columns
                    .Select(column => column < row.Cells.Count && selected.Contains(row.Cells[column]) ? row.Cells[column].Text : string.Empty)
                    .ToList())
                .ToList());
        }

        // A block as a spreadsheet reads it: cells quoted as ForSingleCell quotes them, tabs between
        // cells, "\r\n" between rows - and no line break after the last, for ForSingleCell's reason.
        public static string ForBlock(IReadOnlyList<IReadOnlyList<string>> rows)
        {
            ArgumentNullException.ThrowIfNull(rows);

            return string.Join("\r\n", rows.Select(row => string.Join('\t', row.Select(BoardTableClipboard.ForSingleCell))));
        }

        // One quoted cell starting at `start`, unquoted, and where the text goes on after it - or
        // false when the quotes do not close the cell the way a spreadsheet writes one.
        private static bool TryReadQuoted(string text, int start, out string cell, out int next)
        {
            var value = new StringBuilder();

            for (int at = start + 1; at < text.Length; at++)
            {
                if (text[at] != '"')
                {
                    value.Append(text[at]);
                    continue;
                }

                if (at + 1 < text.Length && text[at + 1] == '"')
                {
                    value.Append('"');
                    at++;
                    continue;
                }

                int after = at + 1;

                if (after == text.Length || text[after] is '\t' or '\r' or '\n')
                {
                    cell = value.ToString();
                    next = after;
                    return true;
                }

                break;
            }

            cell = string.Empty;
            next = start;
            return false;
        }

        // ###########################################################################################
        // What to put in a cell for the given clipboard text. Returns false - and an empty string -
        // when the clipboard holds more than one cell, or nothing at all.
        //
        // The trailing line break Excel adds is dropped; the result is otherwise the cell's own text
        // (BoardTableCell trims it on the way in, as the reader does).
        // ###########################################################################################
        public static bool TryReadSingleCell(string? clipboardText, out string cellText)
        {
            cellText = string.Empty;

            if (string.IsNullOrEmpty(clipboardText))
            {
                return false;
            }

            string text = BoardTableClipboard.WithoutOneTrailingLineBreak(clipboardText);

            // A quoted cell: one value that may legitimately contain line breaks, tabs or quotes.
            if (BoardTableClipboard.TryUnquote(text, out string unquoted))
            {
                cellText = unquoted;
                return true;
            }

            if (text.IndexOfAny(['\t', '\r', '\n']) >= 0)
            {
                return false;
            }

            cellText = text;
            return true;
        }

        // ###########################################################################################
        // The clipboard text for one cell, quoted exactly when Excel would quote it, so pasting it
        // into Excel lands it in one cell rather than splitting it at a line break.
        //
        // No trailing line break: pasting into another program's text field (a browser, an email)
        // is at least as common as pasting into Excel, and a stray newline there is a nuisance.
        // Excel accepts the text without one.
        // ###########################################################################################
        public static string ForSingleCell(string? cellText)
        {
            string text = cellText ?? string.Empty;

            if (text.IndexOfAny(['\t', '\r', '\n', '"']) < 0)
            {
                return text;
            }

            return "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
        }

        private static string WithoutOneTrailingLineBreak(string text)
        {
            if (text.EndsWith("\r\n", StringComparison.Ordinal))
            {
                return text[..^2];
            }

            if (text.EndsWith('\n') || text.EndsWith('\r'))
            {
                return text[..^1];
            }

            return text;
        }

        // ###########################################################################################
        // Unquotes "..." when the WHOLE text is one quoted cell: it starts and ends with a quote,
        // and every quote in between is doubled. "a"<tab>"b" starts and ends with a quote too, but
        // the lone quotes in the middle show it is two cells, so it is not unquoted.
        // ###########################################################################################
        private static bool TryUnquote(string text, out string unquoted)
        {
            unquoted = string.Empty;

            if (text.Length < 2 || text[0] != '"' || text[^1] != '"')
            {
                return false;
            }

            string inner = text[1..^1];

            for (int i = 0; i < inner.Length; i++)
            {
                if (inner[i] != '"')
                {
                    continue;
                }

                if (i + 1 < inner.Length && inner[i + 1] == '"')
                {
                    i++;
                    continue;
                }

                return false;
            }

            unquoted = inner.Replace("\"\"", "\"", StringComparison.Ordinal);
            return true;
        }
    }
}
