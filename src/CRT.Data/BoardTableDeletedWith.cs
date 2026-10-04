using System;
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
        // Whether this is several components deleted together ("U8 and U9") - see Combine.
        public bool IsSeveral { get; init; }

        // ###########################################################################################
        // What went with SEVERAL components deleted at once (owner request, 2026-10-02: "mark
        // multiple rows and then delete those in one go") - one sentence for all of them: the
        // labels joined, each sheet's rows added up in the order the sheets were met, and each
        // component's highlights counted once. A component can be met twice - two regional
        // variants deleted together - and only the first time finds rows still there to take, but
        // both times count its highlights, so those go by label. Null for nothing; one part is
        // returned as it is, so deleting one component reads exactly as it always did.
        // ###########################################################################################
        public static BoardTableDeletedWith? Combine(IReadOnlyList<BoardTableDeletedWith> parts)
        {
            ArgumentNullException.ThrowIfNull(parts);

            if (parts.Count <= 1)
            {
                return parts.Count == 0 ? null : parts[0];
            }

            List<string> labels = parts.Select(part => part.BoardLabel).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            var rows = new List<BoardTableSheetCount>();

            foreach (BoardTableSheetCount count in parts.SelectMany(part => part.Rows))
            {
                int index = rows.FindIndex(existing => existing.Sheet == count.Sheet);

                if (index < 0)
                {
                    rows.Add(count);
                }
                else
                {
                    rows[index] = rows[index] with { Count = rows[index].Count + count.Count };
                }
            }

            int highlights = parts
                .GroupBy(part => part.BoardLabel, StringComparer.OrdinalIgnoreCase)
                .Sum(group => group.Max(part => part.Highlights));

            return new BoardTableDeletedWith(BoardTableDeletedWith.Join(labels), rows, highlights) { IsSeveral = labels.Count > 1 };
        }

        // "Also deleted with U8: 3 rows on Component images and 1 on Component links. Its 2
        // highlights on the schematics go when you save. Undo (Ctrl+Z) brings it all back."
        // Several components: "Also deleted with U8 and U9: ... Their 3 highlights ...".
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
                string highlights = BoardTableDeletedWith.Count(this.Highlights, "highlight", "highlights");
                string verb = this.Highlights == 1 ? "goes" : "go";

                sentences.Add((this.Rows.Count > 0, this.IsSeveral) switch
                {
                    (true, false) => $"Its {highlights} on the schematics {verb} when you save.",
                    (true, true) => $"Their {highlights} on the schematics {verb} when you save.",
                    (false, false) => $"{this.BoardLabel}'s {highlights} on the schematics {verb} when you save.",
                    (false, true) => $"The {highlights} of {this.BoardLabel} on the schematics {verb} when you save."
                });
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
