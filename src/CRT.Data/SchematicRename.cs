using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What renaming a schematic image did - DraftWorkbookStore.RenameSchematic's answer.
    // ###########################################################################################
    public enum SchematicRenameOutcome
    {
        Renamed,

        // The name asked for is the name it has - nothing written.
        Unchanged,

        // A blank name - refused, nothing written.
        Blank,

        // Another image of the board already has the name (whatever its capitals) - refused.
        NameTaken,

        // The draft has no image of the old name (changed outside the window) - nothing written.
        NotFound,

        NoDraft,
        WriteFailed,
    }

    // ###########################################################################################
    // RENAMING A SCHEMATIC IMAGE (owner request, 2026-10-10: "add management of the name of the
    // schematic image also" - the "Schematic name" column of the "Board schematics" sheet).
    //
    // *** THE NAME IS A KEY, SO A RENAME CARRIES WHAT HANGS ON IT. *** It is the sheet's natural
    // key (BoardDraftNaturalKeys), and the board's JSON sidecar files two more things under it:
    // every component highlight drawn on the image, and the image's KiCad calibration. Renaming the
    // row alone would leave the highlights pointing at a schematic that is gone - an error the
    // table and the server both refuse ("highlight.unknown_schematic") - and the next draft save
    // would drop the calibration outright, since it collects calibrations by the board's schematic
    // names (DraftBoardSource.CollectCalibrations). So the rename moves both with the row.
    //
    // *** UNLESS THE OLD NAME STAYS ON ANOTHER ROW. *** Two rows sharing a name is an error the
    // checks report, and renaming one of them is how it gets fixed - but then nothing says which of
    // the two the highlights and the calibration were drawn for, so they stay where they are, on
    // the name that is still there.
    //
    // Names compare trimmed and without regard to capitals, as the checks and the Schematics tab
    // compare them; a change of capitals alone is a rename like any other.
    // ###########################################################################################
    public static class SchematicRename
    {
        public const string NeedsAName = "A schematic image needs a name.";

        public static string NameTaken(string name) =>
            $"Another schematic image is already named [{name}] - each name can be used once.";

        // Why `newName` cannot be an image's name beside `otherNames` (the board's other images),
        // or null when it can.
        public static string? Problem(string? newName, IEnumerable<string?> otherNames)
        {
            string wanted = newName?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
            {
                return SchematicRename.NeedsAName;
            }

            return otherNames.Any(other => SchematicRename.SameName(other, wanted))
                ? SchematicRename.NameTaken(wanted)
                : null;
        }

        // ###########################################################################################
        // The board with the image at `index` renamed, and whether what hangs on its name moved
        // with it: every highlight drawn on it here, and - by DraftWorkbookStore.RenameSchematic,
        // told the same answer - its calibration. Not moved when ANOTHER image keeps the old name (a
        // change of capitals alone is not that - it moves). The board unchanged, and false, when
        // there is no such row.
        //
        // By index, not by name (code review, 2026-10-10): of two images sharing a name, the first
        // was renamed whichever one was meant - SchematicRows finds the row.
        // ###########################################################################################
        public static (BoardData Board, bool MovedItsData) Apply(BoardData board, int index, string newName)
        {
            if (index < 0 || index >= board.Schematics.Count)
            {
                return (board, false);
            }

            string oldName = board.Schematics[index].SchematicName ?? string.Empty;
            string wanted = newName.Trim();

            var schematics = new List<BoardSchematicEntry>(board.Schematics);
            schematics[index] = schematics[index].WithSchematicName(wanted);

            bool oldNameStays = board.Schematics
                .Where((_, other) => other != index)
                .Any(entry => SchematicRename.SameName(entry.SchematicName, oldName));

            if (oldNameStays)
            {
                return (board.WithSchematics(schematics), false);
            }

            List<ComponentHighlightEntry> highlights = board.ComponentHighlights
                .Select(highlight => SchematicRename.SameName(highlight.SchematicName, oldName)
                    ? highlight.WithSchematicName(wanted)
                    : highlight)
                .ToList();

            return (board.WithSchematics(schematics).WithComponentHighlights(highlights), true);
        }

        // ###########################################################################################
        // What hangs on a schematic name, after the schematics were renamed or deleted. `follow` is
        // old name -> new name, or -> null for a schematic deleted (SubmissionRowsBoard
        // .FollowSchematics' map, the server's own rule for a table edit): a highlight or calibration
        // on a renamed schematic moves to the new name, one on a deleted schematic goes, and one on
        // a name `follow` does not list stays as it is.
        // ###########################################################################################
        public static List<ComponentHighlightEntry> FollowHighlights(
            IEnumerable<ComponentHighlightEntry> highlights,
            IReadOnlyDictionary<string, string?> follow) =>
            highlights
                .Select(highlight => SubmissionRowsBoard.Follow(follow, highlight.SchematicName) switch
                {
                    null => null,
                    string name when name == highlight.SchematicName => highlight,
                    string name => highlight.WithSchematicName(name),
                })
                .OfType<ComponentHighlightEntry>()
                .ToList();

        public static List<KiCadCalibrationEntry> FollowCalibrations(
            IEnumerable<KiCadCalibrationEntry> calibrations,
            IReadOnlyDictionary<string, string?> follow) =>
            calibrations
                .Select(calibration => SubmissionRowsBoard.Follow(follow, calibration.SchematicName) switch
                {
                    null => null,
                    string name when name == calibration.SchematicName => calibration,
                    string name => calibration.WithSchematicName(name),
                })
                .OfType<KiCadCalibrationEntry>()
                .ToList();

        // The renames in a `follow` map, as (old name, new name) - the deleted ones left out.
        public static List<(string OldName, string NewName)> RenamesIn(IReadOnlyDictionary<string, string?> follow) =>
            follow
                .Where(pair => pair.Value is not null && !string.Equals(pair.Key, pair.Value, StringComparison.Ordinal))
                .Select(pair => (pair.Key, pair.Value!))
                .ToList();

        public static bool SameName(string? first, string? second) =>
            string.Equals(first?.Trim(), second?.Trim(), StringComparison.OrdinalIgnoreCase);
    }
}
