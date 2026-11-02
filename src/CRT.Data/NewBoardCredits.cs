using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // What a new board's credits were offered: the board as it now reads, every row offered so far, and
    // whether a row followed the board away.
    public sealed record NewBoardCreditOffer(BoardData Board, IReadOnlyList<string> Offered, bool RowsWent = false);

    // ###########################################################################################
    // THE ROWS A NEW BOARD'S CREDITS START WITH (owner request, 2026-11-02: "if this is a new board then
    // make sure to have these rows in 'What': Each individual schematic image; Board data; 'Oscilloscope
    // baseline' for 'What' and 'PAL' or 'NTSC' for 'Sub-category'. Then the user can delete it if needed
    // or he/she should input the missing contact details").
    //
    //   - A ROW NOBODY HAS NAMED IS AN ERROR, ON PURPOSE (BoardRowRequired): the contributor fills in who
    //     did it or deletes the row, and Submit sends nothing until then.
    //   - A ROW IS OFFERED ONCE. What was offered is kept in the draft's marker (DraftMarker
    //     .CreditsOffered), so one deleted stays deleted - also after CRT is closed. A row the credits
    //     have already, typed by hand, counts as offered.
    //   - A ROW NOTHING WAS TYPED IN FOLLOWS THE BOARD: an image renamed or deleted, or the board's
    //     regions changed, takes its row along, and the one wanted now is offered. One with a name in it
    //     is the contributor's, and stays.
    //   - The words are the live boards' own (measured 2026-11-02: "Schematic image" with the image's
    //     name, "Board data", "Oscilloscope baseline" with PAL or NTSC) - no board's own spelling of them
    //     is looked up, as a new board has none yet.
    //
    // PURE: a board and what was offered in, a board and what was offered out. The same board when there
    // is nothing to do. Where a row goes: with the others of its kind, the schematic images first.
    // ###########################################################################################
    public static class NewBoardCredits
    {
        public const string SchematicImage = CreditSuggestions.SchematicImage;

        public const string BoardData = CreditSuggestions.BoardData;

        public const string OscilloscopeBaseline = "Oscilloscope baseline";

        // Which rows a new board wants as it reads now, in the order they stand - on a board with no
        // regions yet, one oscilloscope row with no sub-category.
        public static List<CreditEntry> Wanted(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            var wanted = new List<CreditEntry>();

            foreach (string name in board.Schematics
                .Select(row => row.SchematicName?.Trim() ?? string.Empty)
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                wanted.Add(new CreditEntry { Category = NewBoardCredits.SchematicImage, SubCategory = name });
            }

            wanted.Add(new CreditEntry { Category = NewBoardCredits.BoardData });

            if (BoardRegions.Has(board))
            {
                wanted.AddRange(BoardRegions.All.Select(region => new CreditEntry { Category = NewBoardCredits.OscilloscopeBaseline, SubCategory = region }));
            }
            else
            {
                wanted.Add(new CreditEntry { Category = NewBoardCredits.OscilloscopeBaseline });
            }

            return wanted;
        }

        // What a row is known by among those offered: what it is for and which part of it, capitals aside.
        public static string KeyOf(CreditEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return $"{row.Category?.Trim()}|{row.SubCategory?.Trim()}".ToUpperInvariant();
        }

        // ###########################################################################################
        // `board` with what was offered and is wanted no more gone, if nothing was typed in it, and what
        // is wanted and was never offered added. `Board` is the same object when no row changed.
        // ###########################################################################################
        public static NewBoardCreditOffer Offer(BoardData board, IReadOnlyCollection<string>? offered)
        {
            ArgumentNullException.ThrowIfNull(board);

            List<CreditEntry> wanted = NewBoardCredits.Wanted(board);
            var wantedKeys = wanted.Select(NewBoardCredits.KeyOf).ToHashSet(StringComparer.Ordinal);
            var offeredKeys = (offered ?? []).ToHashSet(StringComparer.Ordinal);

            List<CreditEntry> rows = board.Credits
                .Where(row => !(offeredKeys.Contains(NewBoardCredits.KeyOf(row)) && !wantedKeys.Contains(NewBoardCredits.KeyOf(row)) && NewBoardCredits.NothingTyped(row)))
                .ToList();

            bool rowsWent = rows.Count != board.Credits.Count;
            bool changed = rowsWent;

            // A row that went stops counting as offered: wanted again later, it comes back.
            offeredKeys.IntersectWith(wantedKeys.Concat(rows.Select(NewBoardCredits.KeyOf)));

            var present = rows.Select(NewBoardCredits.KeyOf).ToHashSet(StringComparer.Ordinal);

            foreach (CreditEntry row in wanted)
            {
                string key = NewBoardCredits.KeyOf(row);

                if (!offeredKeys.Add(key) || present.Contains(key))
                {
                    continue;
                }

                rows.Insert(NewBoardCredits.PlaceFor(rows, row), row);
                present.Add(key);
                changed = true;
            }

            return new NewBoardCreditOffer(
                changed ? board.WithBoardSections(credits: rows) : board,
                offeredKeys.Order(StringComparer.Ordinal).ToList(),
                rowsWent);
        }

        // How many rows an offer would add now - what "Board status" says, leaving the draft as it is.
        public static int RowsToAdd(BoardData board, IReadOnlyCollection<string>? offered)
        {
            ArgumentNullException.ThrowIfNull(board);

            var present = board.Credits.Select(NewBoardCredits.KeyOf).ToHashSet(StringComparer.Ordinal);

            return NewBoardCredits.Offer(board, offered).Board.Credits.Count(row => !present.Contains(NewBoardCredits.KeyOf(row)));
        }

        private static bool NothingTyped(CreditEntry row) =>
            string.IsNullOrWhiteSpace(row.NameOrHandle) && string.IsNullOrWhiteSpace(row.Contact);

        // After the last row of its kind - else a schematic image first, as they come first, and the
        // others last.
        private static int PlaceFor(List<CreditEntry> rows, CreditEntry row)
        {
            int last = rows.FindLastIndex(other => string.Equals(other.Category?.Trim(), row.Category, StringComparison.OrdinalIgnoreCase));

            return last >= 0
                ? last + 1
                : string.Equals(row.Category, NewBoardCredits.SchematicImage, StringComparison.Ordinal) ? 0 : rows.Count;
        }
    }
}
