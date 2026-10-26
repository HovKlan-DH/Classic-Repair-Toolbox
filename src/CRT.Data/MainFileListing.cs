using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE DROP-DOWN LISTS IN THE CRT 4.0 MAIN FILE - MasterListing's edits, on
    // "Classic-Repair-Toolbox.v4.0.0.json" (MainDocument) instead of the main Excel data file.
    //
    // MasterListing hands a main file here (IsMainFile), so every caller - a new board's publish,
    // its promotion, a push-back, a deletion, the Boards screen's order - edits the 4.0 lists through
    // the calls it always made. The rules are MasterListing's own: a row is a BOARD (matched by its
    // id, case-insensitively), an insert for a listed board updates its row where it is, a name pair
    // another board is listed under is refused, a row placed after a board no longer listed is
    // refused rather than guessed at, and an edit that changes nothing writes nothing.
    //
    // A row's ExcelDataFile is the board's board file, "<board id>/board.json"
    // (DataTreeLayout.BoardFilePath) - the identity CRT 4.0 carries for a board.
    //
    // *** EVERY WRITE DATES BOTH LISTS *** (owner request, 2026-10-04, for the Excel file's two
    // sheets): the boards' and the oscilloscopes' revision dates are the day of the write. The
    // oscilloscopes and every field this CRT does not know are written back as they were. The file
    // is replaced whole (FileReplacer), so a reader never sees half a list.
    // ###########################################################################################
    internal static class MainFileListing
    {
        public static bool IsMainFile(string? path) =>
            string.Equals(Path.GetExtension(path ?? string.Empty), ".json", StringComparison.OrdinalIgnoreCase);

        public static bool TryRead(string mainPath, out IReadOnlyList<MasterListingRow> rows, out string why)
        {
            rows = [];

            if (!MainFileListing.TryOpen(mainPath, out MainDocument? document, out why))
                return false;

            rows = MainFileListing.Rows(document!);
            return true;
        }

        public static MasterListingEdit Insert(string mainPath, MasterListingRow row, string? afterExcelDataFile, DateTimeOffset? nowUtc)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (!MasterListing.IsWritableRow(row, out string rowProblem))
                return MasterListingEdit.Failed(rowProblem);

            if (!MainFileListing.TryOpen(mainPath, out MainDocument? document, out string why))
                return MasterListingEdit.Failed($"The main data file could not be opened: {why}");

            IReadOnlyList<MasterListingRow> listed = MainFileListing.Rows(document!);

            if (MasterListing.NamesTakenBy(listed, row.BoardId, row.HardwareName, row.BoardName) is MasterListingRow taken)
                return MasterListingEdit.Failed(MasterListing.NamesTakenMessage(taken));

            var wanted = new MainDocument.BoardListingRow
            {
                BoardId = row.BoardId,
                HardwareName = row.HardwareName.Trim(),
                BoardName = row.BoardName.Trim(),
                HardwareNotes = row.Notes ?? string.Empty,
            };

            int existing = MasterListing.IndexOfBoard(listed, row.BoardId);

            if (existing >= 0)
            {
                MainDocument.BoardListingRow current = document!.Boards[existing];

                if (current.BoardId == wanted.BoardId &&
                    current.HardwareName == wanted.HardwareName &&
                    current.BoardName == wanted.BoardName &&
                    current.HardwareNotes == wanted.HardwareNotes)
                {
                    return MasterListingEdit.Done(changed: false);
                }

                wanted.Unknown = current.Unknown;
                document.Boards[existing] = wanted;

                return MainFileListing.Save(document, mainPath, nowUtc);
            }

            int target = 0;

            if (!string.IsNullOrWhiteSpace(afterExcelDataFile))
            {
                string after = afterExcelDataFile.Trim();
                string afterBoard = BoardDescriptorRules.BoardIdFromExcelDataFile(after);

                int found = listed
                    .Select((item, index) => (item, index))
                    .Where(pair =>
                        string.Equals(pair.item.ExcelDataFile, after, StringComparison.OrdinalIgnoreCase) ||
                        (afterBoard.Length > 0 && string.Equals(pair.item.BoardId, afterBoard, StringComparison.OrdinalIgnoreCase)))
                    .Select(pair => pair.index)
                    .DefaultIfEmpty(-1)
                    .First();

                if (found < 0)
                {
                    return MasterListingEdit.Failed(
                        $"It was placed after [{after}], which is no longer in the list. Place it again in the Boards screen.");
                }

                target = found + 1;
            }

            document!.Boards.Insert(target, wanted);

            return MainFileListing.Save(document, mainPath, nowUtc);
        }

        public static MasterListingEdit Remove(string mainPath, string boardId, DateTimeOffset? nowUtc)
        {
            if (!MainFileListing.TryOpen(mainPath, out MainDocument? document, out string why))
                return MasterListingEdit.Failed($"The main data file could not be opened: {why}");

            int removed = document!.Boards.RemoveAll(item =>
                string.Equals(item.BoardId?.Trim(), boardId?.Trim(), StringComparison.OrdinalIgnoreCase));

            return removed == 0 ? MasterListingEdit.Done(changed: false) : MainFileListing.Save(document, mainPath, nowUtc);
        }

        public static MasterListingEdit Reorder(string mainPath, IReadOnlyList<string> boardIdsInOrder, DateTimeOffset? nowUtc)
        {
            ArgumentNullException.ThrowIfNull(boardIdsInOrder);

            if (!MainFileListing.TryOpen(mainPath, out MainDocument? document, out string why))
                return MasterListingEdit.Failed($"The main data file could not be read: {why}");

            IReadOnlyList<int> arranged = MasterListing.ArrangeAs(MainFileListing.Rows(document!), boardIdsInOrder);

            if (arranged.Select((source, place) => source == place).All(same => same))
                return MasterListingEdit.Done(changed: false);

            List<MainDocument.BoardListingRow> before = document!.Boards;
            document.Boards = arranged.Select(index => before[index]).ToList();

            return MainFileListing.Save(document, mainPath, nowUtc);
        }

        // ###########################################################################################
        // The rows as CRT and the server read them: one per listed board, names trimmed, its hardware
        // name carried down from the row above when its own is blank (the Excel sheet's rule, kept for
        // a hand-edited file). THE one reading of the list - CRT's drop-downs (DataManager) and the
        // server's name checks, placement and Boards screen (code review, 2026-10-26: CRT read the
        // names as they stood, so a blank hardware name was a blank entry there and the row above's
        // here). A row with no board id is here too, its ExcelDataFile "".
        // ###########################################################################################
        internal static IReadOnlyList<MasterListingRow> Rows(MainDocument document)
        {
            ArgumentNullException.ThrowIfNull(document);

            var rows = new List<MasterListingRow>(document.Boards.Count);
            string lastHardware = string.Empty;

            foreach (MainDocument.BoardListingRow item in document.Boards)
            {
                string hardware = item.HardwareName?.Trim() ?? string.Empty;
                string id = item.BoardId?.Trim().Trim('/') ?? string.Empty;

                if (hardware.Length > 0)
                    lastHardware = hardware;
                else
                    hardware = lastHardware;

                rows.Add(new MasterListingRow(
                    hardware,
                    item.BoardName?.Trim() ?? string.Empty,
                    id.Length == 0 ? string.Empty : DataTreeLayout.BoardFilePath(id),
                    item.HardwareNotes ?? string.Empty));
            }

            return rows;
        }

        private static bool TryOpen(string mainPath, out MainDocument? document, out string why)
        {
            document = null;
            why = string.Empty;

            if (string.IsNullOrWhiteSpace(mainPath) || !File.Exists(mainPath))
            {
                why = "there is no such file.";
                return false;
            }

            try
            {
                document = DataTreeJson.ReadMain(File.ReadAllText(mainPath, Encoding.UTF8));
                return true;
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
            {
                why = ex.Message;
                return false;
            }
        }

        private static MasterListingEdit Save(MainDocument document, string mainPath, DateTimeOffset? nowUtc)
        {
            string date = BoardWorkbookStyle.FormatRevisionDate(nowUtc ?? DateTimeOffset.UtcNow);

            document.BoardsRevisionDate = date;
            document.OscilloscopesRevisionDate = date;

            byte[] bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(DataTreeJson.Write(document));

            try
            {
                FileReplacer.Replace(mainPath, temporary => File.WriteAllBytes(temporary, bytes));
                return MasterListingEdit.Done(changed: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return MasterListingEdit.Failed($"The main data file could not be written: {ex.Message}");
            }
        }
    }
}
