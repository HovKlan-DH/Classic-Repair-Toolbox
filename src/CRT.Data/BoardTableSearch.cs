using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE BOARD TABLE'S SEARCH BOX (owner request, 2026-10-02: "how about the same search/filter as
    // we have in the "Workbooks" tab, where it highlights what I search for"). Which rows a search
    // shows, and which runs of a cell's text it found; the editor only paints them.
    //
    // *** THE WORKBOOKS TAB'S OWN GRAMMAR *** - WorklogSearchQuery, the one copy: words ANDed (in any
    // cells of the row, each within one cell), "quotes" for a phrase, -minus to leave out, any
    // case. Unlike the Workbooks tab, EVERY cell is searched, numbers included - in a board table
    // "4164" or "100nF" is exactly what one looks for. The +/~/-/! sign in front of a row and the
    // column names are not cells, so they are not searched. A red (deleted) row is searched by the
    // values it shows.
    //
    // *** DECIDED ONCE, WHEN THE SEARCH IS TYPED (agreed case C, 2026-10-02). *** The rows that do
    // not match at that moment are the ones hidden; a row edited afterwards so it no longer matches
    // stays on screen until the search is changed, rather than vanishing under the contributor's
    // typing - and so does a row inserted afterwards. A new, still empty row is never hidden, as
    // with the colour key's pills (BoardTableRowFilter).
    //
    // It works alongside the pills: a row shows when both show it, and the tabs of sheets with no
    // such row are hidden (BoardTableDocument.SheetsShown).
    // ###########################################################################################
    public sealed class BoardTableSearch
    {
        private readonly HashSet<BoardTableRow> thisHidden;

        private BoardTableSearch(string text, WorklogSearchQuery query, HashSet<BoardTableRow> hidden)
        {
            this.Text = text;
            this.Query = query;
            this.thisHidden = hidden;
        }

        // No search: every row shows.
        public static BoardTableSearch None { get; } =
            new(string.Empty, WorklogSearchQuery.Parse(null), new HashSet<BoardTableRow>(ReferenceEqualityComparer.Instance));

        public string Text { get; }

        public WorklogSearchQuery Query { get; }

        public bool IsActive => !this.Query.IsEmpty;

        // The search `text` over `document` as it is now. Blank text is no search (None).
        public static BoardTableSearch For(BoardTableDocument document, string? text)
        {
            ArgumentNullException.ThrowIfNull(document);

            WorklogSearchQuery query = WorklogSearchQuery.Parse(text);
            if (query.IsEmpty)
            {
                return BoardTableSearch.None;
            }

            var hidden = new HashSet<BoardTableRow>(ReferenceEqualityComparer.Instance);

            foreach (BoardTableRow row in document.Sheets.SelectMany(sheet => sheet.Rows))
            {
                if (row.State != BoardTableRowState.Blank && !BoardTableSearch.Matches(query, row))
                {
                    hidden.Add(row);
                }
            }

            return new BoardTableSearch(text!, query, hidden);
        }

        // Whether `query` finds `row` - in its cells, the values it shows.
        public static bool Matches(WorklogSearchQuery query, BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(query);
            ArgumentNullException.ThrowIfNull(row);

            return query.Matches(row.Cells.Select(cell => cell.Text));
        }

        public bool Shows(BoardTableRow row) => !this.thisHidden.Contains(row);

        // Where the search's words are in one cell's text - the runs to mark. None with no search.
        public IReadOnlyList<WorklogSearchHit> HitsIn(string? text) => this.Query.FindHits(text);

        // ###########################################################################################
        // The line above the table when this search, with the picked `kinds`, finds nothing in ANY
        // sheet (owner request, 2026-10-03) - the table then keeps only the sheet on screen, empty
        // (BoardTableDocument.SheetsShown), and this says why. Null when something is found, and
        // with no search: picked pills alone showing nothing were not asked about.
        // ###########################################################################################
        public string? NothingFoundLine(BoardTableDocument document, BoardTableRowKinds kinds)
        {
            ArgumentNullException.ThrowIfNull(document);

            if (!this.IsActive || document.HasRowsShownBy(kinds, this))
            {
                return null;
            }

            string text = this.Text.Trim();

            return kinds == BoardTableRowKinds.None
                ? $"Nothing in any sheet matches \"{text}\""
                : $"Nothing in any sheet matches both \"{text}\" and the counts picked";
        }
    }
}
