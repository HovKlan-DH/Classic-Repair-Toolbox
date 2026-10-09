using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // PLACING A NEW BOARD IN CRT's DROP-DOWN LISTS, on the "Boards" screen (owner request,
    // 2026-09-27: "The maintainer should order the new system, so it becomes visible in the right
    // location for the drop-down lists. This must be done before it can be pushed to BETA.").
    //
    // The list is BETA's main Excel data file as CRT reads it; the new board is one row dragged
    // into it, and its place is saved as "after this row" (null: first). Pure, so the rules and the
    // words are tested - the drag itself is Controls/ListRowDrag.
    //
    // *** THE NAME RULE IS THE SERVER's OWN FUNCTION *** (CRT.Data's MasterListing.NamesTakenBy and
    // IsWritableRow), so the screen cannot accept names the server refuses, or refuse names it
    // would take.
    // ###########################################################################################
    public static class BoardPlacementDisplay
    {
        // Where the placement starts: the one saved on the server, else the server's suggestion.
        public static BoardPlacement Starting(UnlistedBoardEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            return entry.Placement ?? entry.Suggested;
        }

        // ###########################################################################################
        // The index the new row starts at among `listed`: straight after the row whose workbook (or
        // board) is `afterExcelDataFile`, first when that is blank - and at the END, with `isGone`
        // true, when that row is no longer listed, so the screen can say it must be placed again
        // rather than quietly showing a place nobody chose.
        // ###########################################################################################
        public static int StartIndex(IReadOnlyList<BoardListingRow> listed, string? afterExcelDataFile, out bool isGone)
        {
            ArgumentNullException.ThrowIfNull(listed);

            isGone = false;

            if (string.IsNullOrWhiteSpace(afterExcelDataFile))
                return 0;

            string after = afterExcelDataFile.Trim();
            string afterBoard = BoardDescriptorRules.BoardIdFromExcelDataFile(after);

            for (int i = 0; i < listed.Count; i++)
            {
                if (string.Equals(listed[i].ExcelDataFile, after, StringComparison.OrdinalIgnoreCase) ||
                    (afterBoard.Length > 0 && string.Equals(listed[i].BoardId, afterBoard, StringComparison.OrdinalIgnoreCase)))
                {
                    return i + 1;
                }
            }

            isGone = true;
            return listed.Count;
        }

        // The placement a position means: the workbook of the row above it, or null when it is first.
        public static string? AfterAt(IReadOnlyList<string> excelDataFilesInOrder, int newIndex)
        {
            ArgumentNullException.ThrowIfNull(excelDataFilesInOrder);

            return newIndex <= 0 || excelDataFilesInOrder.Count == 0
                ? null
                : excelDataFilesInOrder[Math.Min(newIndex, excelDataFilesInOrder.Count) - 1];
        }

        // ###########################################################################################
        // Why these names cannot be saved, or null when they can - blank names, the characters the
        // file cannot hold, and names ANOTHER board is listed under (CRT keys a board by "hardware
        // name|board name", so two rows with the same pair would be one board twice).
        // ###########################################################################################
        public static string? NameProblem(IReadOnlyList<BoardListingRow> listed, string boardId, string? hardwareName, string? boardName, string? notes)
        {
            ArgumentNullException.ThrowIfNull(listed);

            var row = new MasterListingRow(
                hardwareName?.Trim() ?? string.Empty,
                boardName?.Trim() ?? string.Empty,
                boardId + "/x.xlsx",
                notes?.Trim() ?? string.Empty);

            if (!MasterListing.IsWritableRow(row, out string problem))
                return problem;

            MasterListingRow? taken = MasterListing.NamesTakenBy(
                listed.Select(item => new MasterListingRow(item.HardwareName, item.BoardName, item.ExcelDataFile, string.Empty)).ToList(),
                boardId,
                hardwareName,
                boardName);

            return taken is null ? null : MasterListing.NamesTakenMessage(taken);
        }

        // ###########################################################################################
        // Where CRT will show it, from the order with the new row in it - what the maintainer is
        // really choosing. CRT lists each hardware name once, where it FIRST appears, and a
        // hardware's boards in row order (Main.BoardSelection), so a row placed away from its
        // hardware's other boards still shows under that hardware, at the hardware's first place:
        //
        //   Commodore 128 is hardware 4 of 12 in CRT, and 310378 Open128 its board 3 of 3.
        // ###########################################################################################
        public static string WhereInCrt(IReadOnlyList<(string Hardware, string Board)> orderWithNew, int newIndex)
        {
            ArgumentNullException.ThrowIfNull(orderWithNew);

            if (newIndex < 0 || newIndex >= orderWithNew.Count)
                return string.Empty;

            (string hardware, string board) = orderWithNew[newIndex];
            hardware = hardware.Trim();
            board = board.Trim();

            if (hardware.Length == 0 || board.Length == 0)
                return "Give it a hardware name and a board name to see where CRT shows it.";

            List<string> hardwareNames = orderWithNew
                .Select(item => item.Hardware.Trim())
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            int hardwarePosition = hardwareNames.FindIndex(name => string.Equals(name, hardware, StringComparison.OrdinalIgnoreCase)) + 1;

            List<int> boardsOfHardware = Enumerable.Range(0, orderWithNew.Count)
                .Where(i => string.Equals(orderWithNew[i].Hardware.Trim(), hardware, StringComparison.OrdinalIgnoreCase))
                .ToList();

            int boardPosition = boardsOfHardware.IndexOf(newIndex) + 1;

            return string.Create(
                CultureInfo.InvariantCulture,
                $"{hardware} is hardware {hardwarePosition} of {hardwareNames.Count} in CRT, and {board} its board {boardPosition} of {boardsOfHardware.Count}.");
        }

        // What placing it does, which depends on whether it is in BETA already.
        public static string Explanation(UnlistedBoardEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            return entry.InBeta
                ? "This board is in BETA but not in CRT's drop-down lists, so nobody can choose it. Give it the names CRT shows " +
                  "and drag it to where it belongs - saving adds it to BETA's lists at once."
                : "This new board is not in CRT's drop-down lists yet. Give it the names CRT shows and drag it to where it " +
                  "belongs. It cannot be published to BETA until it has a place, and it is added to the lists when it is.";
        }

        // Whether a place is saved yet - the one thing a maintainer coming back needs to know.
        public static string SavedState(UnlistedBoardEntry entry, bool afterIsGone)
        {
            ArgumentNullException.ThrowIfNull(entry);

            if (entry.Placement is null)
                return "No place saved yet - the one shown is a suggestion.";

            return afterIsGone
                ? "The row it was placed after has left the lists, so it is shown at the end. Place it again."
                : "Its place is saved.";
        }

        public const string CannotPlaceMessage =
            "Only a maintainer of this board or an administrator can place it.";

        // The one-line summary on the chosen board, once it IS listed.
        public static string? ListedLine(BoardListingAnswer? listing, string? boardId)
        {
            BoardListingRow? row = listing?.Rows.FirstOrDefault(item =>
                string.Equals(item.BoardId, boardId, StringComparison.OrdinalIgnoreCase));

            return row is null ? null : $"In CRT's drop-down lists as {row.HardwareName} / {row.BoardName}";
        }

        // The board's entry when it is waiting for a place, or null.
        public static UnlistedBoardEntry? UnlistedEntry(BoardListingAnswer? listing, string? boardId) =>
            listing?.Unlisted.FirstOrDefault(entry =>
                string.Equals(entry.BoardId, boardId, StringComparison.OrdinalIgnoreCase));

        // ###########################################################################################
        // The mark on its row in the Boards list: "Needs a place in the drop-down lists" until one is
        // saved, then "Placed - listed when published to BETA". Null for every other board.
        // ###########################################################################################
        public static string? ListMark(BoardListingAnswer? listing, string? boardId)
        {
            UnlistedBoardEntry? entry = BoardPlacementDisplay.UnlistedEntry(listing, boardId);

            if (entry is null)
                return null;

            return entry.Placement is null
                ? "Needs a place in the drop-down lists"
                : "Placed - listed when published to BETA";
        }

        // ###########################################################################################
        // The line above a submission's table when its board has no place in the lists yet - the
        // approval would be refused, and the maintainer should know before reading the whole table.
        // Null once a place is saved, or for a board already listed.
        // ###########################################################################################
        public static string? ReviewLine(BoardListingAnswer? listing, string? boardId) =>
            BoardPlacementDisplay.UnlistedEntry(listing, boardId) is { Placement: null }
                ? "This new board has no place in CRT's drop-down lists yet, so it cannot be approved. Place it on the Boards screen first."
                : null;

        // Boards needing a place that THIS account can give one - the Boards button's attention.
        public static int NeedingPlace(BoardListingAnswer? listing) =>
            listing?.Unlisted.Count(entry => entry.CanPlace && entry.Placement is null) ?? 0;

        // ###########################################################################################
        // The Boards list's order: the boards waiting for a place first (the ones to act on), then
        // every listed board in CRT's drop-down order, then the rest as the server sent them. With
        // no list known, the server's order unchanged.
        // ###########################################################################################
        public static IReadOnlyList<BoardOverviewEntry> InListOrder(IReadOnlyList<BoardOverviewEntry> boards, BoardListingAnswer? listing)
        {
            ArgumentNullException.ThrowIfNull(boards);

            if (listing is null)
                return boards;

            int Rank(BoardOverviewEntry board)
            {
                int waiting = listing.Unlisted.ToList().FindIndex(entry =>
                    string.Equals(entry.BoardId, board.BoardId, StringComparison.OrdinalIgnoreCase));

                if (waiting >= 0)
                    return waiting;

                int listed = listing.Rows.ToList().FindIndex(row =>
                    string.Equals(row.BoardId, board.BoardId, StringComparison.OrdinalIgnoreCase));

                return listed >= 0 ? listing.Unlisted.Count + listed : int.MaxValue;
            }

            // OrderBy is stable, so the rest keep the server's order among themselves.
            return boards.OrderBy(Rank).ToList();
        }

        // ###########################################################################################
        // After a Save that ran into the two-minute limit (2026-09-28): does the listing, read again,
        // hold the place that was asked for? Either the board is in the list itself (placed straight
        // into BETA's file, its board being there already), or it is still unlisted with exactly this
        // placement kept for its publish. Names are compared as the server keeps them, trimmed.
        // ###########################################################################################
        public static bool HasLanded(SetPlacementRequest request, BoardListingAnswer listing)
        {
            ArgumentNullException.ThrowIfNull(request);
            ArgumentNullException.ThrowIfNull(listing);

            if (listing.Rows.Any(row => string.Equals(row.BoardId, request.BoardId, StringComparison.OrdinalIgnoreCase)))
                return true;

            UnlistedBoardEntry? entry = listing.Unlisted.FirstOrDefault(
                candidate => string.Equals(candidate.BoardId, request.BoardId, StringComparison.OrdinalIgnoreCase));

            return entry?.Placement is BoardPlacement kept &&
                string.Equals(kept.HardwareName, request.HardwareName?.Trim(), StringComparison.Ordinal) &&
                string.Equals(kept.BoardName, request.BoardName?.Trim(), StringComparison.Ordinal) &&
                string.Equals(kept.AfterExcelDataFile ?? string.Empty, request.AfterExcelDataFile ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }
    }
}
