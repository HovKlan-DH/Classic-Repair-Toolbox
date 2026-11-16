using System;
using System.Collections.Generic;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // THE "Oscilloscope" TAB'S OUTPUT BOX KEEPS ITS LAST LinesKept LINES (2026-11-16, performance walkthrough).
    //
    // Every flush (40 ms apart while the scope talks) set the box's whole text again, and the text only ever grew: a
    // TextBox lays all of its text out at every change - measured 2026-10-15, 45 ms at 500 lines - so the longer a
    // session ran, the more each flush cost the UI thread, during keyboard control of the scope, image sync and the
    // full test suite alike. The oldest lines leave the box; every line is in the log as well (AppendOutputLine
    // logs it), so nothing is lost.
    // ###########################################################################################
    internal static class ScopeOutputText
    {
        internal const int LinesKept = 1000;

        // ###########################################################################################
        // `existing` with `lines` appended, each ending in `newLine` - of all of them, the last `kept` lines.
        // `existing` is what the box holds: lines each ending in `newLine`, as this writes them.
        // ###########################################################################################
        public static string Appended(string? existing, IReadOnlyList<string> lines, string newLine, int kept = ScopeOutputText.LinesKept)
        {
            ArgumentException.ThrowIfNullOrEmpty(newLine);

            string text = (existing ?? string.Empty) + string.Join(newLine, lines) + (lines.Count > 0 ? newLine : string.Empty);

            if (kept <= 0)
            {
                return string.Empty;
            }

            // Back from the end: the start of the kept-th last line. The text ends in a line end, which is not a line's start.
            int found = 0;
            int searchEnd = text.Length - newLine.Length;

            while (searchEnd > 0)
            {
                int at = text.LastIndexOf(newLine, searchEnd - 1, StringComparison.Ordinal);

                if (at < 0)
                {
                    break;
                }

                found++;

                if (found == kept)
                {
                    return text[(at + newLine.Length)..];
                }

                searchEnd = at;
            }

            return text;
        }
    }
}
