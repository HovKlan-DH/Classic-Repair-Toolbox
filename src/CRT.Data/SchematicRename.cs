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

    // What a rename decided (SchematicRename.Plan): the outcome, and for Renamed the board renamed,
    // the one rename for the draft's write to follow, and the names from and to.
    public sealed record SchematicRenamePlan(
        SchematicRenameOutcome Outcome,
        BoardData? Board,
        IReadOnlyDictionary<string, string?> Follow,
        string From,
        string To)
    {
        public static SchematicRenamePlan Refused(SchematicRenameOutcome outcome) =>
            new(outcome, null, new Dictionary<string, string?>(), string.Empty, string.Empty);
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
        // Said on the Contribute tab's picture cards (step 3 - the "Schematic images" window went), so
        // in its word: a picture.
        public const string NeedsAName = "A picture needs a name.";

        public static string NameTaken(string name) =>
            $"Another picture is already named [{name}] - each name can be used once.";

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
        // THE WHOLE RENAME DECIDED ON A BOARD - the row found by its name AND its image file
        // (SchematicRows), refused when blank or taken, nothing when the name is the one it has - and
        // the board renamed with `Follow`, the one rename to hand the draft's write (what hangs on the
        // old name moves only when it moved here, see Apply). Board is null for every outcome but
        // Renamed. Shared by the draft file's write (DraftWorkbookStore.RenameSchematic) and the
        // draft held in memory ("Schematic images" through DraftSession, 2026-10-25).
        // ###########################################################################################
        public static SchematicRenamePlan Plan(BoardData current, SchematicRowRef row, string? newName)
        {
            ArgumentNullException.ThrowIfNull(current);
            ArgumentNullException.ThrowIfNull(row);

            string wanted = newName?.Trim() ?? string.Empty;
            int index = SchematicRows.IndexOf(current.Schematics, row);

            if (index < 0)
            {
                return SchematicRenamePlan.Refused(SchematicRenameOutcome.NotFound);
            }

            if (string.Equals(current.Schematics[index].SchematicName?.Trim(), wanted, StringComparison.Ordinal))
            {
                return SchematicRenamePlan.Refused(SchematicRenameOutcome.Unchanged);
            }

            if (SchematicRename.Problem(wanted, current.Schematics.Where((_, other) => other != index).Select(entry => entry.SchematicName)) is not null)
            {
                return SchematicRenamePlan.Refused(wanted.Length == 0 ? SchematicRenameOutcome.Blank : SchematicRenameOutcome.NameTaken);
            }

            string renamedFrom = current.Schematics[index].SchematicName ?? string.Empty;
            (BoardData renamed, bool moved) = SchematicRename.Apply(current, index, wanted);

            var follow = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

            if (moved)
            {
                follow[renamedFrom] = wanted;
            }

            return new SchematicRenamePlan(SchematicRenameOutcome.Renamed, renamed, follow, renamedFrom, wanted);
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
