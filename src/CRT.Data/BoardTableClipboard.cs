using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE CELL TO AND FROM THE CLIPBOARD, the way Excel reads and writes it - for the table
    // editor's Ctrl+C / Ctrl+V on a selected cell (owner request, 2026-09-24).
    //
    // *** ONE CELL AT A TIME, DELIBERATELY. *** The project owner asked for the simple version first:
    // insert a row, then copy cells across one by one. So a clipboard holding a BLOCK of cells
    // (tabs or line breaks - what Excel puts there for more than one cell) is refused rather than
    // squeezed into one cell or spread across several. Pasting a block is a later feature with
    // rules of its own (where it lands, whether it adds rows), not something to half-do here.
    //
    // Excel's clipboard text, which is what these rules follow:
    //   - every copied cell, even a single one, is followed by a line break;
    //   - cells are separated by tabs and rows by line breaks;
    //   - a cell whose text contains a tab, a line break or a double quote is wrapped in double
    //     quotes, with any quote inside doubled ("" for ").
    // ###########################################################################################
    public static class BoardTableClipboard
    {
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
