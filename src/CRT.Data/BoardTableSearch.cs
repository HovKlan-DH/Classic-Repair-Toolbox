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

        // The same, judged only by the cells a sheet's "Pictures" part shows (BoardTableSheetParts): a row
        // listed there for a word in a column it hides would show nothing that says why, and one left out for
        // such a word would have nothing to be left out for (code review, 2026-11-04). It shows among the
        // baselines too, every cell in sight.
        private readonly HashSet<BoardTableRow> thisHiddenAmongPictures;

        // The columns of each sheet it looks in - null for every one.
        private readonly Func<BoardTableSheet, IReadOnlyCollection<int>?>? thisColumnsOf;

        private BoardTableSearch(
            string text,
            WorklogSearchQuery query,
            HashSet<BoardTableRow> hidden,
            HashSet<BoardTableRow> hiddenAmongPictures,
            Func<BoardTableSheet, IReadOnlyCollection<int>?>? columnsOf = null)
        {
            this.Text = text;
            this.Query = query;
            this.thisHidden = hidden;
            this.thisHiddenAmongPictures = hiddenAmongPictures;
            this.thisColumnsOf = columnsOf;
        }

        // No search: every row shows.
        public static BoardTableSearch None { get; } = new(
            string.Empty,
            WorklogSearchQuery.Parse(null),
            new HashSet<BoardTableRow>(ReferenceEqualityComparer.Instance),
            new HashSet<BoardTableRow>(ReferenceEqualityComparer.Instance));

        public string Text { get; }

        public WorklogSearchQuery Query { get; }

        public bool IsActive => !this.Query.IsEmpty;

        // The search `text` over `document` as it is now. Blank text is no search (None).
        public static BoardTableSearch For(BoardTableDocument document, string? text) =>
            BoardTableSearch.Build(document, text, WorklogSearchQuery.Parse(text));

        // ###########################################################################################
        // What "Replace" finds, over `document` as it is now: `text` AS TYPED, capitals ignored - the rows holding it
        // shown, it marked (owner request, 2026-11-07: "I want that 'Find' works in same way as the filtering, where
        // it will filter and highlight my search string (however, do not allow the fancy search - only search for the
        // string as-is ...)"; WorklogSearchQuery.Literal). Empty text is no search.
        // ###########################################################################################
        //
        // *** IN THE COLUMNS TICKED (owner request, 2026-11-07: "I need some possibility to restrict replacement to
        // specific columns, and not necessarily all"; "Columns" in the replace bar). *** `columnsOf` names the columns of
        // a sheet it looks in - null, or a sheet it answers null for, every one: a row is shown for the text in one of
        // THOSE, and only there is it marked (Searches), as only there is it replaced.
        // ###########################################################################################
        public static BoardTableSearch Literal(
            BoardTableDocument document,
            string? text,
            Func<BoardTableSheet, IReadOnlyCollection<int>?>? columnsOf = null) =>
            BoardTableSearch.Build(document, text, WorklogSearchQuery.Literal(text), columnsOf);

        private static BoardTableSearch Build(
            BoardTableDocument document,
            string? text,
            WorklogSearchQuery query,
            Func<BoardTableSheet, IReadOnlyCollection<int>?>? columnsOf = null)
        {
            ArgumentNullException.ThrowIfNull(document);

            if (query.IsEmpty)
            {
                return BoardTableSearch.None;
            }

            var hidden = new HashSet<BoardTableRow>(ReferenceEqualityComparer.Instance);
            var hiddenAmongPictures = new HashSet<BoardTableRow>(ReferenceEqualityComparer.Instance);

            foreach (BoardTableSheet sheet in document.Sheets)
            {
                IReadOnlyCollection<int>? columns = columnsOf?.Invoke(sheet);

                foreach (BoardTableRow row in sheet.Rows)
                {
                    if (row.State == BoardTableRowState.Blank)
                    {
                        continue;
                    }

                    if (!query.Matches(row.Cells.Where(cell => columns is null || columns.Contains(cell.ColumnIndex)).Select(cell => cell.Text)))
                    {
                        hidden.Add(row);
                    }

                    if (BoardTableSheetParts.IsSplit(sheet.Name) &&
                        !query.Matches(BoardTableSheetParts.CellsShown(BoardTableSheetPart.Pictures, row)
                            .Where(cell => columns is null || columns.Contains(cell.ColumnIndex))
                            .Select(cell => cell.Text)))
                    {
                        hiddenAmongPictures.Add(row);
                    }
                }
            }

            return new BoardTableSearch(text!, query, hidden, hiddenAmongPictures, columnsOf);
        }

        // Whether it looks in `column` of `sheet` - every column but those a "Columns" choice left out.
        public bool Searches(BoardTableSheet sheet, int column)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            return this.thisColumnsOf?.Invoke(sheet) is not { } columns || columns.Contains(column);
        }

        // Whether `query` finds `row` - in its cells, the values it shows.
        public static bool Matches(WorklogSearchQuery query, BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(query);
            ArgumentNullException.ThrowIfNull(row);

            return query.Matches(row.Cells.Select(cell => cell.Text));
        }

        public bool Shows(BoardTableRow row) => !this.thisHidden.Contains(row);

        public bool Shows(BoardTableRow row, BoardTableSheetPart part) =>
            !(part == BoardTableSheetPart.Pictures ? this.thisHiddenAmongPictures : this.thisHidden).Contains(row);

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
