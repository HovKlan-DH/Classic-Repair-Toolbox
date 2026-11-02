using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH REGIONS A BOARD COMES IN (owner decisions, 2026-11-01: "PAL and NTSC, or none", stored and
    // published). THE ONE RULE for whether a board shows its PAL/NTSC buttons - under the component
    // list, on the component popup, in the Contribute tab - and offers "Applies to" on its rows.
    //
    // A board says so in the board file's "regions" (BoardData.Regions). One without it is worked out
    // from its rows, as CRT always did: any row saying PAL or NTSC. That rule asked only the COMPONENT
    // rows until 2026-11-01, so a board whose scope baselines were per region and whose components
    // were not hid the buttons and showed one region's baselines for good. Every published board has a
    // regional component, so nothing there needs converting.
    //
    // The switch is all or nothing: 12 of the 14 boards have exactly PAL and NTSC, two have neither.
    // ###########################################################################################
    public static class BoardRegions
    {
        public const string Pal = "PAL";

        public const string Ntsc = "NTSC";

        // The regions a board can come in, as the buttons and the "regions" field say them.
        public static IReadOnlyList<string> All { get; } = [BoardRegions.Pal, BoardRegions.Ntsc];

        public static bool Has(BoardData? board) =>
            board is not null && (board.Regions.Count > 0 || BoardRegions.RegionalRows(board) > 0);

        // ###########################################################################################
        // How many rows of the board say PAL or NTSC - what stops the setting being taken away: a row
        // only one region shows would go on being there with no button to reach it.
        // ###########################################################################################
        public static int RegionalRows(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.Components.Count(row => BoardRegions.IsRegion(row.Region)) +
                   board.ComponentImages.Count(row => BoardRegions.IsRegion(row.Region)) +
                   board.ComponentPins.Count(row => BoardRegions.IsRegion(row.Region)) +
                   board.ComponentReplacements.Count(row => BoardRegions.IsRegion(row.Region)) +
                   board.ComponentLocalFiles.Count(row => BoardRegions.IsRegion(row.Region)) +
                   board.ComponentLinks.Count(row => BoardRegions.IsRegion(row.Region)) +
                   board.BoardLocalFiles.Count(row => BoardRegions.IsRegion(row.Region)) +
                   board.BoardLinks.Count(row => BoardRegions.IsRegion(row.Region));
        }

        // Whether two boards differ in what they say about their regions - what a draft has to submit.
        public static bool Differ(BoardData? published, BoardData? draft) =>
            BoardRegions.Has(published) != BoardRegions.Has(draft);

        // ###########################################################################################
        // The board set to come in PAL and NTSC, or not. The same board when it already says so - and
        // when it is told not to while a row still says PAL or NTSC.
        // ###########################################################################################
        public static BoardData WithRegions(BoardData board, bool has)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (has)
            {
                return board.Regions.Count > 0 ? board : board.WithRegions([.. BoardRegions.All]);
            }

            return board.Regions.Count == 0 || BoardRegions.RegionalRows(board) > 0 ? board : board.WithRegions([]);
        }

        // What a change of the setting is called - "Review changes", the review and the board's history.
        public static string ChangedLine(bool hasNow) =>
            hasNow ? "This board now comes in PAL and NTSC versions" : "This board no longer comes in PAL and NTSC versions";

        // The same of a new board that has them - the review and the board's history (code review, 2026-11-02).
        public const string NewBoardLine = "This new board comes in PAL and NTSC versions";

        // Whether a word is one of the regions - capitals ignored, as every region is compared.
        public static bool IsRegion(string? region)
        {
            string text = region?.Trim() ?? string.Empty;

            return string.Equals(text, BoardRegions.Pal, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, BoardRegions.Ntsc, StringComparison.OrdinalIgnoreCase);
        }
    }
}
