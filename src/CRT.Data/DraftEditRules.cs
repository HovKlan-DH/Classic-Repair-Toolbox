using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT FOLLOWS AN EDIT OF A DRAFT'S BOARD - the rules every draft write applies between the
    // board before an edit and after it. Moved out of DraftWorkbookStore.EditCore (2026-10-25) so the
    // draft held in memory (DraftSession, CRT 4.0 step 1) applies the very same ones: two copies of
    // these would drift, and a highlight left on an old name is an error the checks refuse at submit.
    //
    //   - WHAT HANGS ON A SCHEMATIC'S NAME FOLLOWS IT (owner request, 2026-10-10). The highlights and
    //     the KiCad calibration are filed under a schematic's NAME. A rename moves them, a deletion
    //     drops them - the server's own rule for a table edit (SubmissionRowsBoard.FollowSchematics:
    //     a removed name and an added one on the same image file are a rename). An edit that knows
    //     its rename outright (DraftWorkbookStore.RenameSchematic) hands it in (`follow`).
    //   - WHAT HANGS ON A COMPONENT'S LABEL FOLLOWS IT (code review, 2026-10-13): a label fixed alone
    //     takes its highlights along, by BoardDataDiffer's pairing (ComponentRelabel).
    //   - THE CALIBRATIONS KEPT are those filed under a name the board had before, moved or dropped
    //     with their schematic, plus any filed under a name the edit brought in.
    // ###########################################################################################
    public static class DraftEditRules
    {
        // ###########################################################################################
        // `after` with its highlights following its schematics and labels, and the calibrations to
        // keep - from `calibrations`, the calibrations filed for the board (any filed under neither
        // a name it had nor one it has now are dropped). `follow` is the edit's own schematic
        // renames when it knows them; otherwise they are worked out from the two boards.
        // ###########################################################################################
        public static (BoardData Board, List<KiCadCalibrationEntry> Calibrations) Follow(
            BoardData before,
            BoardData after,
            IEnumerable<KiCadCalibrationEntry> calibrations,
            IReadOnlyDictionary<string, string?>? follow = null)
        {
            ArgumentNullException.ThrowIfNull(before);
            ArgumentNullException.ThrowIfNull(after);
            ArgumentNullException.ThrowIfNull(calibrations);

            // The very same list is no rename and no relabel - skipped, not worked out (code review,
            // 2026-10-26: the table hands over the sheets it changed, the rest as the draft has them).
            follow ??= ReferenceEquals(before.Schematics, after.Schematics)
                ? new Dictionary<string, string?>()
                : SubmissionRowsBoard.FollowSchematics(before.Schematics, after.Schematics);

            BoardData updated = after;

            if (follow.Count > 0)
            {
                updated = updated.WithComponentHighlights(SchematicRename.FollowHighlights(updated.ComponentHighlights, follow));
            }

            IReadOnlyDictionary<string, string> relabels = ReferenceEquals(before.Components, updated.Components)
                ? new Dictionary<string, string>()
                : ComponentRelabel.Between(before.Components, updated.Components);

            if (relabels.Count > 0)
            {
                updated = updated.WithComponentHighlights(ComponentRelabel.FollowHighlights(updated.ComponentHighlights, relabels));
            }

            return (updated, DraftEditRules.CalibrationsAfter(before, after, calibrations, follow));
        }

        // The schematic names a board had before an edit, then those it has after that it did not -
        // what DraftWorkbookStore reads the calibrations of from a board's file.
        public static IEnumerable<string> NamesBeforeAndAfter(BoardData before, BoardData after)
        {
            HashSet<string> had = DraftEditRules.NamesOf(before);

            return had.Concat(DraftEditRules.NamesOf(after).Where(name => !had.Contains(name)));
        }

        private static List<KiCadCalibrationEntry> CalibrationsAfter(
            BoardData before,
            BoardData after,
            IEnumerable<KiCadCalibrationEntry> calibrations,
            IReadOnlyDictionary<string, string?> follow)
        {
            HashSet<string> had = DraftEditRules.NamesOf(before);
            HashSet<string> has = DraftEditRules.NamesOf(after);
            List<KiCadCalibrationEntry> filed = calibrations.ToList();

            List<KiCadCalibrationEntry> kept = SchematicRename.FollowCalibrations(
                filed.Where(entry => had.Contains(entry.SchematicName)).OrderBy(entry => entry.SchematicName, StringComparer.OrdinalIgnoreCase),
                follow);

            var named = new HashSet<string>(kept.Select(entry => entry.SchematicName.Trim()), StringComparer.OrdinalIgnoreCase);

            kept.AddRange(filed
                .Where(entry => has.Contains(entry.SchematicName) && !named.Contains(entry.SchematicName))
                .OrderBy(entry => entry.SchematicName, StringComparer.OrdinalIgnoreCase));

            return kept;
        }

        // ###########################################################################################
        // What an edit took off the board along with schematic images it removed: the images' names
        // and what was drawn on them, as `before` had it - their highlights and KiCad calibrations,
        // which Follow drops with them. A rename is no removal. Said to the contributor afterwards
        // (BoardTableSessionWording.WentWith), with Ctrl+Z to bring it back.
        // ###########################################################################################
        public static DraftEditLoss LostWith(BoardContent before, BoardContent after)
        {
            ArgumentNullException.ThrowIfNull(before);
            ArgumentNullException.ThrowIfNull(after);

            List<string> removed = SubmissionRowsBoard.FollowSchematics(before.Board.Schematics, after.Board.Schematics)
                .Where(pair => pair.Value is null && pair.Key.Trim().Length > 0)
                .Select(pair => pair.Key.Trim())
                .ToList();

            if (removed.Count == 0)
            {
                return DraftEditLoss.None;
            }

            var names = new HashSet<string>(removed, StringComparer.OrdinalIgnoreCase);

            return new DraftEditLoss(
                removed,
                before.Board.ComponentHighlights.Count(highlight => names.Contains(highlight.SchematicName?.Trim() ?? string.Empty)),
                before.Calibrations.Count(calibration => names.Contains(calibration.SchematicName?.Trim() ?? string.Empty)));
        }

        private static HashSet<string> NamesOf(BoardData board) => new(
            board.Schematics.Select(entry => entry.SchematicName?.Trim() ?? string.Empty).Where(name => name.Length > 0),
            StringComparer.OrdinalIgnoreCase);
    }

    // What an edit took off the board with the schematic images it removed - see DraftEditRules.LostWith.
    public sealed record DraftEditLoss(IReadOnlyList<string> Schematics, int Highlights, int Calibrations)
    {
        public static DraftEditLoss None { get; } = new([], 0, 0);
    }
}
