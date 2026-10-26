using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HOW A BOARD'S SCOPE BASELINES WERE MEASURED (CRT 4.0 step 2, 2026-10-26 - the plan: "One board
    // field holds how the measurements were made"; owner decision the same day: "Do what you
    // recommend" - the board files the data already keeps, no new field).
    //
    // The live trees already say it: five boards carry text files in the board-file category
    // "Scope measurements" ("Introduction to C64 scope measurements", under "Scope baseline/"). So a
    // board's measuring notes ARE its board files of that category - edited under Files and links like
    // any other, travelling and published as they always have, with no field, no wire change and no
    // server change. The component pane's scope baselines and the popup's link to them.
    // ###########################################################################################
    public static class ScopeMeasurementFiles
    {
        // The board-file category the notes are kept under, as the live trees spell it.
        public const string Category = "Scope measurements";

        // The board's files of that category - capitals and spaces aside - in the board's order, each
        // with a file to open.
        public static List<BoardLocalFileEntry> Of(BoardData? board) => ScopeMeasurementFiles.Among(board?.BoardLocalFiles);

        // The same, among a board's list of files.
        public static List<BoardLocalFileEntry> Among(IEnumerable<BoardLocalFileEntry>? boardFiles) =>
            boardFiles?
                .Where(row => string.Equals(row.Category?.Trim(), ScopeMeasurementFiles.Category, StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(row.File))
                .ToList()
            ?? [];
    }
}
