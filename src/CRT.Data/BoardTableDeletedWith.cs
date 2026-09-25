using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What went with a component deleted in the board table (owner request, 2026-09-25): its
    // rows on the other component sheets, deleted at once, and its highlights, which go when the
    // table is saved. See BoardTableDocument.DeleteRowsOfComponent.
    //
    // The sentence is built here, not in the editor, so the Drafts tab's table and the review
    // application's say the same thing.
    // ###########################################################################################
    public sealed record BoardTableDeletedWith(
        string BoardLabel,
        IReadOnlyList<BoardTableSheetCount> Rows,
        int Highlights)
    {
        // "Also deleted with U8: 3 rows on Component images and 1 on Component links. Its 2
        // highlights on the schematics go when you save. Undo (Ctrl+Z) brings it all back."
        public string Describe()
        {
            var sentences = new List<string>();

            if (this.Rows.Count > 0)
            {
                IEnumerable<string> parts = this.Rows.Select((sheet, i) => i == 0
                    ? $"{BoardTableDeletedWith.Count(sheet.Count, "row", "rows")} on {sheet.Sheet}"
                    : $"{sheet.Count} on {sheet.Sheet}");

                sentences.Add($"Also deleted with {this.BoardLabel}: {BoardTableDeletedWith.Join(parts.ToList())}.");
            }

            if (this.Highlights > 0)
            {
                string owner = this.Rows.Count > 0 ? "Its" : $"{this.BoardLabel}'s";
                string verb = this.Highlights == 1 ? "goes" : "go";

                sentences.Add($"{owner} {BoardTableDeletedWith.Count(this.Highlights, "highlight", "highlights")} on the schematics {verb} when you save.");
            }

            sentences.Add("Undo (Ctrl+Z) brings it all back.");

            return string.Join(" ", sentences);
        }

        private static string Count(int count, string one, string many) => $"{count} {(count == 1 ? one : many)}";

        // "a", "a and b", "a, b and c".
        private static string Join(IReadOnlyList<string> parts) =>
            parts.Count <= 1
                ? string.Join(string.Empty, parts)
                : $"{string.Join(", ", parts.Take(parts.Count - 1))} and {parts[^1]}";
    }

    // How many rows went from one sheet.
    public sealed record BoardTableSheetCount(string Sheet, int Count);
}
