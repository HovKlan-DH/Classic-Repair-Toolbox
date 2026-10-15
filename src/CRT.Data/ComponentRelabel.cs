using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A COMPONENT RELABELLED IS FOLLOWED BY ITS HIGHLIGHTS (code review, 2026-10-13).
    //
    // A schematic renamed takes its highlights and calibration with it (SchematicRename); a
    // component relabelled did not. Highlights are filed by label in the sidecar, so a typo fixed in
    // the Components sheet ("U8" -> "U08") left the rectangles on "U8": a warning on the schematic, a
    // component with no highlight, and on every user's machine after publishing, a stray rectangle
    // and an unmarked chip.
    //
    // The rule is BoardDataDiffer's, the one the table and the server count with (owner decision,
    // 2026-10-04): a row whose KEY alone changed is that row renamed (PairRenamedRows); a key AND
    // another cell changed is a new row and a deleted one, and nothing is followed. A label is moved
    // only when no component still carries the old one (a regional twin keeps it) and the rename is
    // to ONE new label.
    //
    // Used by every draft write (DraftWorkbookStore.EditCore) and by a maintainer's amendment on
    // the server (ReviewTableContract.WithTableSections) - one rule on both sides.
    // ###########################################################################################
    public static class ComponentRelabel
    {
        // Old label -> new label, for the components `after` relabelled from `before`.
        public static IReadOnlyDictionary<string, string> Between(
            IReadOnlyList<ComponentEntry>? before,
            IReadOnlyList<ComponentEntry>? after)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (before is null || after is null)
                return map;

            var beforeKeys = new HashSet<string>(before.Select(BoardDraftNaturalKeys.ForRow), StringComparer.OrdinalIgnoreCase);
            var afterKeys = new HashSet<string>(after.Select(BoardDraftNaturalKeys.ForRow), StringComparer.OrdinalIgnoreCase);

            List<ComponentEntry> removed = before.Where(row => !afterKeys.Contains(BoardDraftNaturalKeys.ForRow(row))).ToList();
            List<ComponentEntry> added = after.Where(row => !beforeKeys.Contains(BoardDraftNaturalKeys.ForRow(row))).ToList();

            if (removed.Count == 0 || added.Count == 0)
                return map;

            var labelsAfter = new HashSet<string>(after.Select(row => Label(row)), StringComparer.OrdinalIgnoreCase);
            var candidates = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardRowRename rename in BoardDataDiffer.PairRenamedRows(removed.Cast<object>().ToList(), added.Cast<object>().ToList()))
            {
                string from = Label(removed[rename.Removed]);
                string to = Label(added[rename.Added]);

                if (from.Length == 0 || to.Length == 0 || string.Equals(from, to, StringComparison.OrdinalIgnoreCase))
                    continue; // A region-only or capitalisation-only edit: the label still matches.

                if (labelsAfter.Contains(from))
                    continue; // Another component still carries the old label.

                if (!candidates.TryGetValue(from, out HashSet<string>? targets))
                    candidates[from] = targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                targets.Add(to);
            }

            foreach ((string from, HashSet<string> targets) in candidates)
            {
                if (targets.Count == 1)
                    map[from] = targets.First();
            }

            return map;
        }

        // The highlights with every relabelled component's under its new label.
        public static List<ComponentHighlightEntry> FollowHighlights(
            IEnumerable<ComponentHighlightEntry> highlights,
            IReadOnlyDictionary<string, string> relabels)
        {
            ArgumentNullException.ThrowIfNull(highlights);

            if (relabels is null || relabels.Count == 0)
                return highlights.ToList();

            return highlights
                .Select(highlight => relabels.TryGetValue((highlight.BoardLabel ?? string.Empty).Trim(), out string? to)
                    ? highlight.WithBoardLabel(to)
                    : highlight)
                .ToList();
        }

        private static string Label(ComponentEntry row) => (row.BoardLabel ?? string.Empty).Trim();
    }
}
