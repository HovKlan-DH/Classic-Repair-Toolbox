using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHETHER A DRAFT HOLDS ITS ROWS IN ANOTHER ORDER THAN THE PUBLISHED BOARD - and in which lists.
    //
    // *** ORDER IS A CHANGE (owner decision, 2026-10-30: "a change of (any) order should also be
    // submitted as a change"). *** BoardDataDiffer counts rows - an added, removed or changed one -
    // and a row only MOVED is none of those, so the order is counted here, beside it, by everything
    // that asks whether a draft holds work: the draft's row and its Submit (DraftOtherChanges), the
    // board parts' "Draft" mark, "Review changes" (DraftChangeSummary), retirement, and the summary
    // a maintainer reads (ReviewSummary.Compare - the server's, sent with the submission). The server
    // publishes a submission's rows as sent, so the order reaches the published board with it.
    //
    // The lists a person orders: every sheet of the board (ReviewSummary.AllSections) but the
    // components' highlights, whose order nobody sees - a rectangle drawn again in another place of
    // the list is no work. Rows are told apart by BoardDraftNaturalKeys.ForRow, the pairing rule every
    // comparison uses; only the rows BOTH boards have are compared - an added or removed row is the
    // differ's to count, and moves nothing else.
    // ###########################################################################################
    public static class BoardRowOrder
    {
        private static readonly IReadOnlyList<ReviewSection> Sheets = ReviewSummary.AllSections
            .Where(section => !string.Equals(section.Name, ReviewSummary.SectionComponentHighlights, StringComparison.Ordinal))
            .ToList();

        public static bool DiffersFrom(BoardData published, BoardData draft) =>
            BoardRowOrder.ReorderedSheets(published, draft).Count > 0;

        // ###########################################################################################
        // The sheets whose shared rows stand in another order in `draft` than in `published` - by their
        // sheet names (BoardWorkbookSchema), in the board's own order of sheets.
        // ###########################################################################################
        public static IReadOnlyList<string> ReorderedSheets(BoardData published, BoardData draft)
        {
            ArgumentNullException.ThrowIfNull(published);
            ArgumentNullException.ThrowIfNull(draft);

            var reordered = new List<string>();

            foreach (ReviewSection sheet in BoardRowOrder.Sheets)
            {
                if (BoardRowOrder.Reordered(sheet.RowsOf(published), sheet.RowsOf(draft)))
                {
                    reordered.Add(sheet.Name);
                }
            }

            return reordered;
        }

        // ###########################################################################################
        // THE COMPONENTS WHOSE OWN ROWS STAND IN ANOTHER ORDER (code review, 2026-11-02: U8's links dragged
        // read as unchanged in the pane, with Submit and "Review changes" saying the order had changed).
        // Each component's rows are compared among themselves, list by list, so a component only moved past
        // another - a category dragged, every chip of it along - is none: that is the list's order, said for
        // the board. By its label as the draft spells it, trimmed.
        // ###########################################################################################
        public static IReadOnlySet<string> ComponentsReordered(BoardData published, BoardData draft)
        {
            ArgumentNullException.ThrowIfNull(published);
            ArgumentNullException.ThrowIfNull(draft);

            return new HashSet<string>(
                BoardRowOrder.ReorderedByLabel(published.Components, draft.Components, row => row.BoardLabel)
                    .Concat(BoardRowOrder.ReorderedByLabel(published.ComponentImages, draft.ComponentImages, row => row.BoardLabel))
                    .Concat(BoardRowOrder.ReorderedByLabel(published.ComponentLocalFiles, draft.ComponentLocalFiles, row => row.BoardLabel))
                    .Concat(BoardRowOrder.ReorderedByLabel(published.ComponentLinks, draft.ComponentLinks, row => row.BoardLabel))
                    .Concat(BoardRowOrder.ReorderedByLabel(published.ComponentPins, draft.ComponentPins, row => row.BoardLabel))
                    .Concat(BoardRowOrder.ReorderedByLabel(published.ComponentReplacements, draft.ComponentReplacements, row => row.BoardLabel)),
                StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The same answer for ONE component (code review, 2026-11-02): the component pane asks it on the
        // UI thread after each component shown and each change, and grouping every component of the
        // board to answer for one is work nobody asked for.
        // ###########################################################################################
        public static bool ComponentReordered(BoardData published, BoardData draft, string? label)
        {
            ArgumentNullException.ThrowIfNull(published);
            ArgumentNullException.ThrowIfNull(draft);

            string wanted = label?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
            {
                return false;
            }

            bool OwnReordered<T>(IEnumerable<T> before, IEnumerable<T> after, Func<T, string?> labelOf)
                where T : class
            {
                bool Own(T row) => string.Equals(labelOf(row)?.Trim(), wanted, StringComparison.OrdinalIgnoreCase);

                return BoardRowOrder.Reordered(before.Where(Own), after.Where(Own));
            }

            return OwnReordered(published.Components, draft.Components, row => row.BoardLabel) ||
                   OwnReordered(published.ComponentImages, draft.ComponentImages, row => row.BoardLabel) ||
                   OwnReordered(published.ComponentLocalFiles, draft.ComponentLocalFiles, row => row.BoardLabel) ||
                   OwnReordered(published.ComponentLinks, draft.ComponentLinks, row => row.BoardLabel) ||
                   OwnReordered(published.ComponentPins, draft.ComponentPins, row => row.BoardLabel) ||
                   OwnReordered(published.ComponentReplacements, draft.ComponentReplacements, row => row.BoardLabel);
        }

        private static IEnumerable<string> ReorderedByLabel<T>(IEnumerable<T> published, IEnumerable<T> draft, Func<T, string?> labelOf)
            where T : class
        {
            ILookup<string, T> before = published.ToLookup(row => labelOf(row)?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase);

            return draft
                .GroupBy(row => labelOf(row)?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .Where(rows => rows.Key.Length > 0 && BoardRowOrder.Reordered(before[rows.Key], rows))
                .Select(rows => rows.Key);
        }

        // Whether one list's shared rows stand in another order - the board parts' "Draft" mark asks
        // it of a single list.
        public static bool Reordered<T>(IEnumerable<T>? published, IEnumerable<T>? draft)
            where T : class =>
            BoardRowOrder.Reordered(published?.Cast<object>(), draft?.Cast<object>());

        private static bool Reordered(IEnumerable<object>? published, IEnumerable<object>? draft)
        {
            if (BoardRowOrder.KeysOf(BoardRowOrder.AsAFileHolds(published)) is not { } before ||
                BoardRowOrder.KeysOf(BoardRowOrder.AsAFileHolds(draft)) is not { } after)
            {
                return false;
            }

            var shared = new HashSet<string>(before, StringComparer.OrdinalIgnoreCase);
            shared.IntersectWith(after);

            return !before.Where(shared.Contains).SequenceEqual(after.Where(shared.Contains), StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // *** THE ORDER A BOARD FILE CAN HOLD (code review, 2026-11-01). *** A file holds a board's
        // pictures first and its scope baselines after them (BoardDocumentMapping), and the draft held
        // in memory has them in one list: a baseline dragged among the pictures in the table was said to
        // be another order, which the save lost and Submit never sent. So each kind is compared among
        // its own, as the file will have it. Every other list is left as it stands.
        // ###########################################################################################
        private static IEnumerable<object>? AsAFileHolds(IEnumerable<object>? rows) =>
            rows?.OrderBy(row => ComponentImageKinds.IsScopeBaseline(row as ComponentImageEntry) ? 1 : 0);

        // ###########################################################################################
        // A list's order changed, in words a person reads - "Order of the files changed" - for the
        // draft's summary and the line above a maintainer's table. A sheet this does not know is named
        // as the sheet.
        // ###########################################################################################
        public static string OrderChangedLine(string sheet) => $"Order of the {BoardRowOrder.ListName(sheet)} changed";

        public static string ListName(string sheet) => sheet switch
        {
            BoardWorkbookSchema.SheetBoardSchematics => "schematic images",
            BoardWorkbookSchema.SheetComponents => "components",
            BoardWorkbookSchema.SheetComponentImages => "component pictures",
            BoardWorkbookSchema.SheetComponentLocalFiles => "component files",
            BoardWorkbookSchema.SheetComponentLinks => "component links",
            BoardWorkbookSchema.SheetComponentPins => "component pins",
            BoardWorkbookSchema.SheetComponentReplacements => "replacement parts",
            BoardWorkbookSchema.SheetBoardLocalFiles => "files",
            BoardWorkbookSchema.SheetBoardLinks => "links",
            BoardWorkbookSchema.SheetCredits => "credits",
            BoardWorkbookSchema.SheetKiCadImportantSignals => "important KiCad signals",
            _ => sheet,
        };

        // A list's row keys in order, the first of rows sharing a key only - or null for a list that
        // holds no board rows (ForRow knows no such row).
        private static List<string>? KeysOf(IEnumerable? rows)
        {
            var keys = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (object? row in rows ?? Array.Empty<object>())
            {
                if (row is null)
                {
                    continue;
                }

                string key;

                try
                {
                    key = BoardDraftNaturalKeys.ForRow(row);
                }
                catch (NotSupportedException)
                {
                    return null;
                }

                if (seen.Add(key))
                {
                    keys.Add(key);
                }
            }

            return keys;
        }
    }
}
