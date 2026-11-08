using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE "REPLACE ALL" REPLACES IN ONE CELL'S TEXT - the one rule BoardTableDocument.ReplaceAll writes by, the count
    // on its button shows and the table's preview draws (owner request, 2026-11-07: "when I put some text in 'Replace
    // with' then it will overwrite the highlighted text with another color and then the content of this text, so I can
    // see directly what the new text will look like" - and "a counter badge for how many places it will replace").
    //
    // Every occurrence of the text found, capitals ignored, left to right and never overlapping - as String.Replace
    // with OrdinalIgnoreCase replaces: "aaa" holds "aa" once. The search box's own marks merge overlapping finds, which
    // would mark more than a replace touches, so the preview never reads them.
    // ###########################################################################################
    public static class BoardTableReplaceText
    {
        // How often `find` occurs in `text`.
        public static int Occurrences(string? text, string? find)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(find))
            {
                return 0;
            }

            int count = 0;

            for (int at = text.IndexOf(find, StringComparison.OrdinalIgnoreCase); at >= 0; at = text.IndexOf(find, at + find.Length, StringComparison.OrdinalIgnoreCase))
            {
                count++;
            }

            return count;
        }

        // ###########################################################################################
        // `text` cut into what stays and what is replaced, in order, covering all of it - the found pieces as they are
        // written in the cell. No find: the text whole, unmatched; empty text: nothing.
        // ###########################################################################################
        public static IReadOnlyList<(string Text, bool IsFound)> Split(string? text, string? find)
        {
            var parts = new List<(string Text, bool IsFound)>();

            if (string.IsNullOrEmpty(text))
            {
                return parts;
            }

            if (string.IsNullOrEmpty(find))
            {
                parts.Add((text, false));
                return parts;
            }

            int position = 0;

            for (int at = text.IndexOf(find, StringComparison.OrdinalIgnoreCase); at >= 0; at = text.IndexOf(find, at + find.Length, StringComparison.OrdinalIgnoreCase))
            {
                if (at > position)
                {
                    parts.Add((text[position..at], false));
                }

                parts.Add((text.Substring(at, find.Length), true));
                position = at + find.Length;
            }

            if (position < text.Length)
            {
                parts.Add((text[position..], false));
            }

            return parts;
        }
    }
}
