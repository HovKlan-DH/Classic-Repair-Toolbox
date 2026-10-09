using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Represents a single hardware/board entry from the main Excel data file.
    // Used to populate hardware and board drop-down selectors in the UI.
    // ###########################################################################################
    public class HardwareBoardEntry
    {
        public string HardwareName { get; init; } = string.Empty;
        public string BoardName { get; init; } = string.Empty;
        public string ExcelDataFile { get; init; } = string.Empty;
        public string HardwareNotes { get; init; } = string.Empty;

        // ###########################################################################################
        // True for a board that exists only as a local draft - created through "Add a new board"
        // (NewContributeStrategy.md Phase 2, session 2c, task 9) and never synced, so its
        // ExcelDataFile names a file that does not exist and never will (see NewBoardIdentity).
        //
        // This is a stored flag rather than an inferred File.Exists check because four places have
        // to act on it and inferring would be four copies of the same guess. Two of them matter a
        // great deal: DataManager's sync/orphan bookkeeping must not add a phantom path to the set
        // of board files it expects on disk, and thisTryAddMappedFilesFromBoardExcel FAILS CLOSED
        // when a board Excel cannot be read - so without this flag a single draft-only board would
        // silently disable orphan cleanup for every real board too.
        // ###########################################################################################
        public bool IsDraftOnly { get; init; }

        // ###########################################################################################
        // True for a board listed only by a legacy "_UserContribution" workbook - a board made the
        // old way, whose files sit in Data/ but were never published by anyone.
        // ###########################################################################################
        public bool IsUserContribution { get; init; }

        // ###########################################################################################
        // Whether the MAIN workbook lists this board - whether the file at its path in Data/ is
        // really the published board.
        //
        // *** A _UserContribution BOARD IS NOT (2026-09-27). *** Its file in Data/ is the
        // contributor's own. Draft retirement compared a draft of such a board with that file, found
        // them identical (the draft was a copy of it), and deleted the draft three minutes after the
        // maintainer merged it into BETA - "its changes are now in the published data", which they
        // were not. Anything that asks "is this in the published data" must ask this.
        // ###########################################################################################
        public bool IsPublished => !this.IsDraftOnly && !this.IsUserContribution;

        public override string ToString() => $"{this.HardwareName} - {this.BoardName}";

        // ###########################################################################################
        // Short "hardware/board" label derived from ExcelDataFile's own folder structure, e.g.
        // "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx" -> "C64/250407" - the immediate parent
        // folder is the board MODEL (e.g. "250407"), and its own parent is the hardware (e.g. "C64").
        // Deliberately not HardwareName/BoardName: those are the full names from the main Excel sheet
        // ("Commodore 64" and "250407 (short board)"), which read well as combo box entries with
        // their own dedicated row each but are too long side by side in a compact label - see the
        // worklog bar's workbook picker, the one place that needed something shorter.
        //
        // ExcelDataFile uses "/" (the sync manifest's own separator - see DataManager), so split on
        // that rather than Path.DirectorySeparatorChar, which is "\" on Windows and would not match.
        // Falls back to ExcelDataFile itself (or "") when the path does not have at least two folder
        // segments to take - malformed data rather than the normal case, but still something readable
        // instead of a blank or a crash.
        // ###########################################################################################
        public string ShortHardwareBoardLabel
        {
            get
            {
                if (string.IsNullOrWhiteSpace(this.ExcelDataFile))
                    return string.Empty;

                var segments = this.ExcelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries);

                // At least .../<hardware>/<board>/<file.xlsx> is needed to take both folder names.
                if (segments.Length < 3)
                    return this.ExcelDataFile;

                string board = segments[^2];
                string hardware = segments[^3];
                return $"{hardware}/{board}";
            }
        }
    }
}