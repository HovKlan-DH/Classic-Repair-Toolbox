using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE MAIN EXCEL DATA FILE'S "Hardware & Board" SHEET - the one definition of its names.
    //
    // CRT reads it to fill the hardware and board drop-downs (DataManager.LoadMainExcel), the
    // server's orphan rule reads it for the boards it lists (DataTreeUsage), and since 2026-09-27 the
    // server WRITES one row into it for a new system (MasterListing). Three readers and a writer of
    // one file format: a column renamed on one side only would be a board silently missing from
    // everyone's lists, so the names live here once.
    // ###########################################################################################
    public static class MasterWorkbookSchema
    {
        public const string SheetName = "Hardware & Board";
        public const string ColHardwareName = "Hardware name in drop-down";
        public const string ColBoardName = "Board name in drop-down";
        public const string ColExcelDataFile = "Excel data file";
        public const string ColHardwareNotes = "Hardware notes in \"Overview\" tab";

        // How far down the sheet the header row is looked for - the sheet opens with a preamble.
        public const int HeaderSearchRows = 20;
    }

    // ###########################################################################################
    // One row of the drop-down lists: what CRT shows (hardware and board names, the Overview notes)
    // and the board workbook it loads. Hardware is always the row's own name - a blank cell carried
    // forward from the row above is resolved when read.
    // ###########################################################################################
    public sealed record MasterListingRow(string HardwareName, string BoardName, string ExcelDataFile, string Notes)
    {
        // "Commodore/C128/310378 Open128" - which system the row is, whatever its file is called.
        public string SystemId => SystemDescriptorRules.SystemIdFromExcelDataFile(this.ExcelDataFile);
    }

    // What an edit did. Changed is false for an edit that found the file already as asked - a re-run.
    public sealed record MasterListingEdit(bool IsDone, bool Changed, string? Failure)
    {
        public static MasterListingEdit Done(bool changed) => new(true, changed, null);

        public static MasterListingEdit Failed(string failure) => new(false, false, failure);
    }

    // ###########################################################################################
    // ADDING A NEW SYSTEM TO THE DROP-DOWN LISTS (owner request, 2026-09-27): "When a system is added
    // to BETA, and it is a NEW system, can you then make sure it gets added also to the main Excel
    // data file in the Data root? The maintainer should order the new system, so it becomes visible
    // in the right location for the drop-down lists."
    //
    // *** ONE ROW, EDITED IN PLACE - NEVER A REGENERATED FILE. *** The file carries a preamble, the
    // Oscilloscope sheet and formatting nobody here models; rewriting it from a model would drop them.
    // So the row is inserted into the existing sheet (with its neighbour's formatting), and nothing
    // else is touched.
    //
    // *** ONLY THE NEWEST GENERATION IS EVER WRITTEN. *** Older masters serve older builds and are
    // frozen for good (NewContributeStrategy.md, open question 5) - NewestMasterPath answers null
    // rather than name one, so a tree with no versioned master is refused, not written.
    //
    // *** A ROW IS A SYSTEM. *** Matched by its system id (its folders), never its display names, and
    // an insert for a system already listed updates that row instead of adding a second - so re-running
    // an interrupted publish is safe, as every other step of a publish is.
    //
    // ATOMIC: written to a temporary file beside it and moved into place, so a crash leaves the old
    // file whole. Order-sensitive readers (the drop-downs) never see half a list.
    // ###########################################################################################
    public static class MasterListing
    {
        // ###########################################################################################
        // The newest-generation master in a data root, or null when there is none to write.
        // ###########################################################################################
        public static string? NewestMasterPath(string? dataRoot)
        {
            Version? generation = DataGenerationRules.ResolveNewestGenerationFromTree(dataRoot);
            if (generation is null)
            {
                return null;
            }

            string path = Path.Combine(dataRoot!, DataGenerationRules.BuildMasterFileName(generation));

            return File.Exists(path) ? path : null;
        }

        // ###########################################################################################
        // The rows, top to bottom, as CRT reads them: a row naming neither a board nor a workbook is
        // skipped, and a blank hardware cell takes the name above it. False, with a reason, when the
        // file or its sheet or header cannot be read - never an empty list, which would read as "this
        // system is not listed" and invite a duplicate row.
        // ###########################################################################################
        public static bool TryRead(string masterPath, out IReadOnlyList<MasterListingRow> rows, out string why)
        {
            rows = [];

            if (!MasterListing.TryOpen(masterPath, out ExcelPackage? package, out why))
            {
                return false;
            }

            using (package)
            {
                if (!MasterListing.TryLocate(package!, requireEveryColumn: false, out SheetLayout? layout, out why))
                {
                    return false;
                }

                rows = MasterListing.ReadRows(layout!).Select(row => row.Row).ToList();
                return true;
            }
        }

        public static int IndexOfSystem(IReadOnlyList<MasterListingRow> rows, string? systemId)
        {
            string id = systemId?.Trim() ?? string.Empty;

            if (id.Length == 0)
            {
                return -1;
            }

            for (int i = 0; i < rows.Count; i++)
            {
                if (string.Equals(rows[i].SystemId, id, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        // ###########################################################################################
        // The OTHER system already listed under these two names, or null. CRT keys a board by
        // "hardware name|board name", case-insensitively (its settings, its workbooks and the
        // drop-downs all do), so two rows with the same pair would be one board twice in the lists
        // with every setting shared between them. The system's own row is not a clash.
        // ###########################################################################################
        public static MasterListingRow? NamesTakenBy(
            IReadOnlyList<MasterListingRow> rows,
            string? systemId,
            string? hardwareName,
            string? boardName)
        {
            string hardware = hardwareName?.Trim() ?? string.Empty;
            string board = boardName?.Trim() ?? string.Empty;
            string id = systemId?.Trim() ?? string.Empty;

            return rows.FirstOrDefault(row =>
                !string.Equals(row.SystemId, id, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(row.HardwareName.Trim(), hardware, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(row.BoardName.Trim(), board, StringComparison.OrdinalIgnoreCase));
        }

        // The sentence for a clash - the same words wherever it is refused.
        public static string NamesTakenMessage(MasterListingRow taken)
        {
            ArgumentNullException.ThrowIfNull(taken);

            return $"[{taken.HardwareName}] / [{taken.BoardName}] is already in the drop-down lists, for {taken.SystemId}. " +
                "Give this system a board name of its own in the Systems screen.";
        }

        // ###########################################################################################
        // Adds the row after the row whose workbook is afterExcelDataFile - or FIRST when that is
        // blank - or, for a system already listed, updates its row where it is.
        //
        // Names another system is listed under are REFUSED (NamesTakenBy) - at the placement, and
        // again here, since a publish or a promotion writes later and the lists can have changed.
        //
        // A named "after" row that is no longer in the list is REFUSED rather than guessed at: the
        // maintainer placed the system relative to it, and anywhere else is a place nobody chose.
        // ###########################################################################################
        public static MasterListingEdit Insert(string masterPath, MasterListingRow row, string? afterExcelDataFile)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (!MasterListing.IsWritableRow(row, out string rowProblem))
            {
                return MasterListingEdit.Failed(rowProblem);
            }

            if (!MasterListing.TryOpen(masterPath, out ExcelPackage? package, out string why))
            {
                return MasterListingEdit.Failed($"The main Excel data file could not be opened: {why}");
            }

            using (package)
            {
                if (!MasterListing.TryLocate(package!, requireEveryColumn: true, out SheetLayout? layout, out why))
                {
                    return MasterListingEdit.Failed($"The main Excel data file cannot take the row: {why}");
                }

                List<ListedRow> listed = MasterListing.ReadRows(layout!);

                if (MasterListing.NamesTakenBy(listed.Select(item => item.Row).ToList(), row.SystemId, row.HardwareName, row.BoardName)
                    is MasterListingRow taken)
                {
                    return MasterListingEdit.Failed(MasterListing.NamesTakenMessage(taken));
                }

                ListedRow? existing = listed.FirstOrDefault(item =>
                    string.Equals(item.Row.SystemId, row.SystemId, StringComparison.OrdinalIgnoreCase));

                if (existing is not null)
                {
                    if (existing.Row == row)
                    {
                        return MasterListingEdit.Done(changed: false);
                    }

                    MasterListing.WriteCells(layout!, existing.SheetRow, row);
                    return MasterListing.Save(package!, masterPath);
                }

                int target;

                if (string.IsNullOrWhiteSpace(afterExcelDataFile))
                {
                    target = layout!.HeaderRow + 1;
                }
                else
                {
                    string afterSystem = SystemDescriptorRules.SystemIdFromExcelDataFile(afterExcelDataFile.Trim());

                    ListedRow? after = listed.FirstOrDefault(item =>
                        string.Equals(item.Row.ExcelDataFile, afterExcelDataFile.Trim(), StringComparison.OrdinalIgnoreCase) ||
                        (afterSystem.Length > 0 && string.Equals(item.Row.SystemId, afterSystem, StringComparison.OrdinalIgnoreCase)));

                    if (after is null)
                    {
                        return MasterListingEdit.Failed(
                            $"It was placed after [{afterExcelDataFile.Trim()}], which is no longer in the list. " +
                            "Place it again in the Systems screen.");
                    }

                    target = after.SheetRow + 1;
                }

                // The row that will sit just below the new one, with the hardware name it has NOW -
                // needed if its own cell is blank, see below.
                ListedRow? below = listed.FirstOrDefault(item => item.SheetRow >= target);

                ExcelWorksheet sheet = layout!.Sheet;
                sheet.InsertRow(target, 1);

                // Formatted like its neighbour: the row above, or for a first row the one now below.
                int styleSource = target - 1 > layout.HeaderRow ? target - 1 : target + 1;
                int lastColumn = Math.Max(sheet.Dimension?.End.Column ?? 0, layout.LastColumn);

                for (int column = 1; column <= lastColumn; column++)
                {
                    sheet.Cells[target, column].StyleID = sheet.Cells[styleSource, column].StyleID;
                }

                MasterListing.WriteCells(layout, target, row);

                // *** A BLANK HARDWARE CELL BELOW MUST NOT INHERIT THIS ROW'S NAME. *** CRT carries
                // a blank hardware cell forward from the row above, so inserting "Commodore 128"
                // above a blank-named "Commodore 64" board would move that board. Its own name is
                // written into its cell instead.
                if (below is not null && string.IsNullOrWhiteSpace(sheet.Cells[below.SheetRow + 1, layout.HardwareColumn].Text))
                {
                    sheet.Cells[below.SheetRow + 1, layout.HardwareColumn].Value = below.Row.HardwareName;
                }

                return MasterListing.Save(package!, masterPath);
            }
        }

        // ###########################################################################################
        // Takes a system's row out - a new system pushed back out of BETA before it ever reached
        // production. Nothing to remove is a success that changed nothing.
        // ###########################################################################################
        public static MasterListingEdit Remove(string masterPath, string systemId)
        {
            if (!MasterListing.TryOpen(masterPath, out ExcelPackage? package, out string why))
            {
                return MasterListingEdit.Failed($"The main Excel data file could not be opened: {why}");
            }

            using (package)
            {
                if (!MasterListing.TryLocate(package!, requireEveryColumn: false, out SheetLayout? layout, out why))
                {
                    return MasterListingEdit.Failed($"The main Excel data file could not be read: {why}");
                }

                List<ListedRow> listed = MasterListing.ReadRows(layout!);

                ListedRow? row = listed.FirstOrDefault(item =>
                    string.Equals(item.Row.SystemId, systemId?.Trim(), StringComparison.OrdinalIgnoreCase));

                if (row is null)
                {
                    return MasterListingEdit.Done(changed: false);
                }

                // The row below keeps its own hardware name if its cell leaned on this one's.
                ListedRow? below = listed.FirstOrDefault(item => item.SheetRow > row.SheetRow);

                if (below is not null &&
                    string.IsNullOrWhiteSpace(layout!.Sheet.Cells[below.SheetRow, layout.HardwareColumn].Text))
                {
                    layout.Sheet.Cells[below.SheetRow, layout.HardwareColumn].Value = below.Row.HardwareName;
                }

                layout!.Sheet.DeleteRow(row.SheetRow, 1);

                return MasterListing.Save(package!, masterPath);
            }
        }

        // ###########################################################################################
        // WHERE A SYSTEM GOES IN ANOTHER TREE'S LIST - production's, when it is promoted from BETA
        // (owner decision, 2026-09-27: "insert at the same place"). After the nearest row ABOVE it in
        // the source that the target also lists; failing that, before the nearest one BELOW it; and
        // when the target lists none of its neighbours at all, first.
        //
        // False when the source does not list the system - nothing to place.
        // ###########################################################################################
        public static bool TryResolvePlacement(
            IReadOnlyList<MasterListingRow> source,
            string systemId,
            IReadOnlyList<MasterListingRow> target,
            out string? afterExcelDataFile)
        {
            afterExcelDataFile = null;

            int index = MasterListing.IndexOfSystem(source, systemId);
            if (index < 0)
            {
                return false;
            }

            for (int i = index - 1; i >= 0; i--)
            {
                int inTarget = MasterListing.IndexOfSystem(target, source[i].SystemId);

                if (inTarget >= 0)
                {
                    afterExcelDataFile = target[inTarget].ExcelDataFile;
                    return true;
                }
            }

            for (int i = index + 1; i < source.Count; i++)
            {
                int inTarget = MasterListing.IndexOfSystem(target, source[i].SystemId);

                if (inTarget >= 0)
                {
                    afterExcelDataFile = inTarget == 0 ? null : target[inTarget - 1].ExcelDataFile;
                    return true;
                }
            }

            return true;
        }

        // ###########################################################################################
        // A row fit to write: every name present, one line, and not absurdly long. Notes may be empty
        // and may run to several lines - the Overview tab shows them as written.
        // ###########################################################################################
        public static bool IsWritableRow(MasterListingRow row, out string problem)
        {
            ArgumentNullException.ThrowIfNull(row);

            foreach ((string value, string label) in new[] { (row.HardwareName, "hardware name"), (row.BoardName, "board name") })
            {
                string trimmed = value?.Trim() ?? string.Empty;

                if (trimmed.Length == 0)
                {
                    problem = $"The {label} cannot be blank.";
                    return false;
                }

                if (trimmed.Length > MasterListing.MaximumNameLength)
                {
                    problem = $"The {label} is longer than {MasterListing.MaximumNameLength} characters.";
                    return false;
                }

                if (trimmed.Any(char.IsControl))
                {
                    problem = $"The {label} must be on one line.";
                    return false;
                }
            }

            if (row.SystemId.Length == 0)
            {
                problem = $"[{row.ExcelDataFile}] does not name a board workbook.";
                return false;
            }

            if ((row.Notes?.Length ?? 0) > MasterListing.MaximumNotesLength)
            {
                problem = $"The notes are longer than {MasterListing.MaximumNotesLength} characters.";
                return false;
            }

            problem = string.Empty;
            return true;
        }

        public const int MaximumNameLength = 100;
        public const int MaximumNotesLength = 2000;

        // ------------------------------------------------------------------ the sheet

        private sealed record SheetLayout(
            ExcelWorksheet Sheet,
            int HeaderRow,
            int HardwareColumn,
            int BoardColumn,
            int ExcelDataFileColumn,
            int NotesColumn,
            int LastColumn);

        private sealed record ListedRow(int SheetRow, MasterListingRow Row);

        private static bool TryOpen(string masterPath, out ExcelPackage? package, out string why)
        {
            package = null;
            why = string.Empty;

            if (string.IsNullOrWhiteSpace(masterPath) || !File.Exists(masterPath))
            {
                why = "there is no such file.";
                return false;
            }

            EpplusLicense.Ensure();

            try
            {
                // Read fully into memory, so the file itself is not held open while it is replaced.
                byte[] bytes = File.ReadAllBytes(masterPath);
                package = new ExcelPackage(new MemoryStream(bytes));
                return true;
            }
            catch (Exception ex)
            {
                package?.Dispose();
                package = null;
                why = ex.Message;
                return false;
            }
        }

        private static bool TryLocate(ExcelPackage package, bool requireEveryColumn, out SheetLayout? layout, out string why)
        {
            layout = null;
            why = string.Empty;

            ExcelWorksheet? sheet;

            try
            {
                sheet = package.Workbook.Worksheets[MasterWorkbookSchema.SheetName];
            }
            catch (Exception ex)
            {
                why = ex.Message;
                return false;
            }

            if (sheet?.Dimension is null)
            {
                why = $"it has no [{MasterWorkbookSchema.SheetName}] sheet.";
                return false;
            }

            int lastRow = sheet.Dimension.End.Row;
            int lastColumn = sheet.Dimension.End.Column;

            for (int row = 1; row <= Math.Min(MasterWorkbookSchema.HeaderSearchRows, lastRow); row++)
            {
                int Find(string header)
                {
                    for (int column = 1; column <= lastColumn; column++)
                    {
                        if (string.Equals(sheet.Cells[row, column].Text?.Trim(), header, StringComparison.OrdinalIgnoreCase))
                        {
                            return column;
                        }
                    }

                    return -1;
                }

                int excelColumn = Find(MasterWorkbookSchema.ColExcelDataFile);
                if (excelColumn < 0)
                {
                    continue;
                }

                int hardware = Find(MasterWorkbookSchema.ColHardwareName);
                int board = Find(MasterWorkbookSchema.ColBoardName);
                int notes = Find(MasterWorkbookSchema.ColHardwareNotes);

                if (requireEveryColumn && (hardware < 0 || board < 0 || notes < 0))
                {
                    why = $"its [{MasterWorkbookSchema.SheetName}] sheet does not have all four columns " +
                          $"([{MasterWorkbookSchema.ColHardwareName}], [{MasterWorkbookSchema.ColBoardName}], " +
                          $"[{MasterWorkbookSchema.ColExcelDataFile}], [{MasterWorkbookSchema.ColHardwareNotes}]).";
                    return false;
                }

                layout = new SheetLayout(sheet, row, hardware, board, excelColumn, notes, lastColumn);
                return true;
            }

            why = $"its [{MasterWorkbookSchema.SheetName}] sheet has no [{MasterWorkbookSchema.ColExcelDataFile}] column.";
            return false;
        }

        private static List<ListedRow> ReadRows(SheetLayout layout)
        {
            var rows = new List<ListedRow>();
            string lastHardware = string.Empty;
            int lastRow = layout.Sheet.Dimension?.End.Row ?? layout.HeaderRow;

            string Cell(int row, int column) =>
                column < 0 ? string.Empty : layout.Sheet.Cells[row, column].Text?.Trim() ?? string.Empty;

            for (int row = layout.HeaderRow + 1; row <= lastRow; row++)
            {
                string hardware = Cell(row, layout.HardwareColumn);
                string board = Cell(row, layout.BoardColumn);
                string excel = DataTreeUsage.Normalise(Cell(row, layout.ExcelDataFileColumn));
                string notes = layout.NotesColumn < 0 ? string.Empty : layout.Sheet.Cells[row, layout.NotesColumn].Text ?? string.Empty;

                if (hardware.Length > 0)
                {
                    lastHardware = hardware;
                }
                else
                {
                    hardware = lastHardware;
                }

                if (board.Length == 0 && excel.Length == 0)
                {
                    continue;
                }

                rows.Add(new ListedRow(row, new MasterListingRow(hardware, board, excel, notes)));
            }

            return rows;
        }

        private static void WriteCells(SheetLayout layout, int row, MasterListingRow values)
        {
            layout.Sheet.Cells[row, layout.HardwareColumn].Value = values.HardwareName.Trim();
            layout.Sheet.Cells[row, layout.BoardColumn].Value = values.BoardName.Trim();
            layout.Sheet.Cells[row, layout.ExcelDataFileColumn].Value = values.ExcelDataFile.Trim();
            layout.Sheet.Cells[row, layout.NotesColumn].Value = values.Notes ?? string.Empty;
        }

        private static MasterListingEdit Save(ExcelPackage package, string masterPath)
        {
            string temporary = masterPath + ".tmp_" + Guid.NewGuid().ToString("N");

            try
            {
                package.SaveAs(new FileInfo(temporary));
                File.Move(temporary, masterPath, overwrite: true);

                return MasterListingEdit.Done(changed: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
            {
                try
                {
                    if (File.Exists(temporary))
                    {
                        File.Delete(temporary);
                    }
                }
                catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException)
                {
                    CrtLog.Warning($"Could not remove [{temporary}] - [{cleanup.Message}]");
                }

                return MasterListingEdit.Failed($"The main Excel data file could not be written: {ex.Message}");
            }
        }
    }
}
