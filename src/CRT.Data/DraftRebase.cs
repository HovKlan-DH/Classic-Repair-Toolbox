using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Handlers.DataHandling
{
    // One place where the contributor's draft and the published update both changed the same
    // thing. The contributor's version is the one kept; this says so.
    public sealed record DraftRebaseConflict(string Section, string Row, string What);

    // What a rebase produces: the draft's board and calibrations as they should now be, what was
    // kept of the contributor's work, and where the two sides met.
    public sealed class DraftRebaseResult
    {
        public required BoardData Board { get; init; }

        public required List<KiCadCalibrationEntry> Calibrations { get; init; }

        // How many of the contributor's own changes (rows added, changed, removed or renamed, and
        // calibrations) were carried onto the published update.
        public int ChangesKept { get; init; }

        public IReadOnlyList<DraftRebaseConflict> Conflicts { get; init; } = [];
    }

    // ###########################################################################################
    // BRINGING A DRAFT UP TO DATE WITH THE PUBLISHED BOARD (code review, 2026-10-13).
    //
    // A draft is a COMPLETE copy of the board (DraftSeeder - "deletion means deletion"), and a
    // submission is the board's complete intended state. So a draft started before somebody else's
    // work was published does not hold that work, and sending it as it was would ask the
    // maintainer to remove it again - every row published since read as a deletion by this
    // contributor. The drift window used to promise "your edits are still applied on top of it",
    // which had stopped being true when the overlay model went (Phase 6).
    //
    // This is the three-way merge that makes it true. BASE is the published board as it was copied
    // into the draft (DraftBase), DRAFT is the contributor's board now, PUBLISHED is the board as it
    // is now. Per section, by the SAME identity the Drafts tab counts with (BoardDraftNaturalKeys,
    // case-insensitive, and BoardDataDiffer.PairRenamedRows for a row whose key alone changed):
    //
    //   - a row the contributor left exactly as copied follows the published update - changed,
    //     deleted or untouched there;
    //   - a row the contributor changed, added, removed or renamed keeps the contributor's version;
    //   - a row the published update added is added.
    //
    // Where BOTH sides changed one row the contributor's version wins - it is their draft, and the
    // maintainer sees it as their change - and the meeting is listed as a conflict, so the
    // contributor can look.
    //
    // *** A ROW THE UPDATE RELABELLED IS RELABELLED ON THE OLD SIDES FIRST (code review,
    // 2026-10-13). *** By the same rule (PairRenamedRows, base -> published), before the merge: the
    // base's row and the contributor's row of the same identity take the update's key, so a
    // description the contributor edited lands on the update's "U08" instead of bringing "U8" back
    // beside it (and the contributor's highlights of a relabelled component follow it -
    // ComponentRelabel). A key the contributor's draft already uses is left alone.
    //
    // DUPLICATES ARE KEPT: a section may hold two rows with one key (a warning in the table, not an
    // error), so each row's identity is its key AND its occurrence among rows with that key - the
    // second "U8" pairs with the second "U8". A highlight's identity is ALL its values (several
    // rectangles may mark one component), so a moved rectangle is the old one removed and the new
    // one added.
    //
    // ORDER follows the published board, with the contributor's added rows placed after the row
    // they followed in the draft. Order is not a change BoardDataDiffer counts (CLAUDE.md), so this
    // only has to be sensible, not exact.
    //
    // Pure: nothing here reads or writes a file (DraftRebaser does).
    // ###########################################################################################
    public static class DraftRebase
    {
        public static DraftRebaseResult Merge(
            BoardData based,
            BoardData draft,
            BoardData published,
            IReadOnlyList<KiCadCalibrationEntry>? baseCalibrations,
            IReadOnlyList<KiCadCalibrationEntry>? draftCalibrations,
            IReadOnlyList<KiCadCalibrationEntry>? publishedCalibrations)
        {
            ArgumentNullException.ThrowIfNull(based);
            ArgumentNullException.ThrowIfNull(draft);
            ArgumentNullException.ThrowIfNull(published);

            var state = new MergeState();

            // ###########################################################################################
            // *** A SCHEMATIC THE UPDATE RENAMED IS RENAMED ON THE OLD SIDES FIRST. *** Highlights and
            // calibrations are filed by schematic NAME, so a rename by the update would otherwise read
            // as "removed" on one side and "added" on the other - and a contributor's recalibration
            // of the old name would land beside the update's calibration of the new one. The update's
            // renames (the server's rule: same image file, new name) are applied to the base and to
            // what the contributor left under the old name before anything is merged. A schematic the
            // contributor renamed has no row under the old name left to apply it to, so theirs stands.
            // ###########################################################################################
            Dictionary<string, string> renamedByUpdate = SubmissionRowsBoard
                .FollowSchematics(based.Schematics, published.Schematics)
                .Where(pair => pair.Value is { Length: > 0 })
                .ToDictionary(pair => pair.Key, pair => pair.Value!, StringComparer.OrdinalIgnoreCase);

            if (renamedByUpdate.Count > 0)
            {
                based = DraftRebase.RenameSchematics(based, renamedByUpdate);
                draft = DraftRebase.RenameSchematics(draft, renamedByUpdate);
                baseCalibrations = DraftRebase.RenameCalibrations(baseCalibrations, renamedByUpdate);
                draftCalibrations = DraftRebase.RenameCalibrations(draftCalibrations, renamedByUpdate);
            }

            // A component the update relabelled takes the highlights the base and the contributor
            // have of it along - the rule every draft write uses (ComponentRelabel).
            IReadOnlyDictionary<string, string> relabelledByUpdate = ComponentRelabel.Between(based.Components, published.Components);

            if (relabelledByUpdate.Count > 0)
            {
                based = based.WithComponentHighlights(ComponentRelabel.FollowHighlights(based.ComponentHighlights, relabelledByUpdate));
                draft = draft.WithComponentHighlights(ComponentRelabel.FollowHighlights(draft.ComponentHighlights, relabelledByUpdate));
            }

            var board = new BoardData
            {
                RevisionDate = published.RevisionDate,
                HardwareName = DraftRebase.Prefer(published.HardwareName, draft.HardwareName),
                BoardName = DraftRebase.Prefer(published.BoardName, draft.BoardName),
                Schematics = DraftRebase.MergeSection("Board schematics", based.Schematics, draft.Schematics, published.Schematics, NaturalKey, state),
                Components = DraftRebase.MergeSection("Components", based.Components, draft.Components, published.Components, NaturalKey, state),
                ComponentImages = DraftRebase.MergeSection("Component images", based.ComponentImages, draft.ComponentImages, published.ComponentImages, NaturalKey, state),
                ComponentHighlights = DraftRebase.MergeSection("Component highlights", based.ComponentHighlights, draft.ComponentHighlights, published.ComponentHighlights, AllValues, state, pairRenames: false),
                ComponentLocalFiles = DraftRebase.MergeSection("Component local files", based.ComponentLocalFiles, draft.ComponentLocalFiles, published.ComponentLocalFiles, NaturalKey, state),
                ComponentLinks = DraftRebase.MergeSection("Component links", based.ComponentLinks, draft.ComponentLinks, published.ComponentLinks, NaturalKey, state),
                BoardLocalFiles = DraftRebase.MergeSection("Board local files", based.BoardLocalFiles, draft.BoardLocalFiles, published.BoardLocalFiles, NaturalKey, state),
                BoardLinks = DraftRebase.MergeSection("Board links", based.BoardLinks, draft.BoardLinks, published.BoardLinks, NaturalKey, state),
                Credits = DraftRebase.MergeSection("Credits", based.Credits, draft.Credits, published.Credits, NaturalKey, state),
                KiCadImportantSignals = DraftRebase.MergeSection("Important signals", based.KiCadImportantSignals, draft.KiCadImportantSignals, published.KiCadImportantSignals, NaturalKey, state),
            };

            List<KiCadCalibrationEntry> calibrations = DraftRebase.MergeSection(
                "KiCad calibrations",
                baseCalibrations ?? [],
                draftCalibrations ?? [],
                publishedCalibrations ?? [],
                CalibrationKey,
                state,
                pairRenames: false);

            // ###########################################################################################
            // *** WHAT HANGS ON A SCHEMATIC'S NAME FOLLOWS THE MERGED SCHEMATICS. *** A schematic the
            // published update renamed (or deleted) still has, in the contributor's own highlights
            // and calibration, its OLD name. The rule every draft write uses (EditCore - the server's
            // SubmissionRowsBoard.FollowSchematics: same image file, new name = a rename) is applied
            // from the BASE's schematics to the merged ones, so nothing is left on a name the board
            // no longer has. Names that are new on either side are not in the map and stay.
            // ###########################################################################################
            IReadOnlyDictionary<string, string?> follow = SubmissionRowsBoard.FollowSchematics(based.Schematics, board.Schematics);

            if (follow.Count > 0)
            {
                board = board.WithComponentHighlights(SchematicRename.FollowHighlights(board.ComponentHighlights, follow));
                calibrations = SchematicRename.FollowCalibrations(calibrations, follow);
            }

            return new DraftRebaseResult
            {
                Board = board,
                Calibrations = calibrations,
                ChangesKept = state.ChangesKept,
                Conflicts = state.Conflicts,
            };
        }

        // The board with every schematic named in `renames` - its row, and the highlights on it -
        // under its new name.
        private static BoardData RenameSchematics(BoardData board, IReadOnlyDictionary<string, string> renames)
        {
            string NewName(string? name) =>
                name is not null && renames.TryGetValue(name.Trim(), out string? renamed) ? renamed : name ?? string.Empty;

            return board
                .WithSchematics(board.Schematics.Select(row => NewName(row.SchematicName) is var name && name != row.SchematicName ? row.WithSchematicName(name) : row).ToList())
                .WithComponentHighlights(board.ComponentHighlights.Select(row => NewName(row.SchematicName) is var name && name != row.SchematicName ? row.WithSchematicName(name) : row).ToList());
        }

        private static IReadOnlyList<KiCadCalibrationEntry> RenameCalibrations(
            IReadOnlyList<KiCadCalibrationEntry>? calibrations,
            IReadOnlyDictionary<string, string> renames) =>
            (calibrations ?? [])
                .Select(entry => entry.SchematicName is not null && renames.TryGetValue(entry.SchematicName.Trim(), out string? renamed)
                    ? entry.WithSchematicName(renamed)
                    : entry)
                .ToList();

        private static string NaturalKey<T>(T row) where T : class => BoardDraftNaturalKeys.ForRow(row);

        private static string AllValues(ComponentHighlightEntry row) => string.Join(
            BoardDraftNaturalKeys.Separator,
            row.SchematicName?.Trim(), row.BoardLabel?.Trim(), row.X?.Trim(), row.Y?.Trim(), row.Width?.Trim(), row.Height?.Trim());

        private static string CalibrationKey(KiCadCalibrationEntry entry) => entry.SchematicName?.Trim() ?? string.Empty;

        // The published value where it has one, the draft's otherwise (a caption the published
        // workbook lost should not blank the draft's).
        private static string Prefer(string? first, string? second) =>
            string.IsNullOrWhiteSpace(first) ? second ?? string.Empty : first;

        private sealed class MergeState
        {
            public int ChangesKept { get; set; }

            public List<DraftRebaseConflict> Conflicts { get; } = [];
        }

        // One row in the result being built: its identity and the row itself.
        private sealed class Slot<T>(string key, T row)
        {
            public string Key { get; set; } = key;

            public T Row { get; set; } = row;
        }

        // ###########################################################################################
        // The three-way merge of one section. See the class header for the rules.
        // ###########################################################################################
        private static List<T> MergeSection<T>(
            string section,
            IReadOnlyList<T>? basedRows,
            IReadOnlyList<T>? draftRows,
            IReadOnlyList<T>? publishedRows,
            Func<T, string> identity,
            MergeState state,
            bool pairRenames = true)
            where T : class
        {
            List<(string Key, T Row)> based = DraftRebase.Keyed(basedRows, identity);
            List<(string Key, T Row)> draft = DraftRebase.Keyed(draftRows, identity);
            List<(string Key, T Row)> published = DraftRebase.Keyed(publishedRows, identity);

            if (pairRenames)
                DraftRebase.FollowUpdateRenames(based, draft, published);

            var baseByKey = based.ToDictionary(entry => entry.Key, entry => entry.Row, StringComparer.OrdinalIgnoreCase);
            var draftByKey = draft.ToDictionary(entry => entry.Key, entry => entry.Row, StringComparer.OrdinalIgnoreCase);
            var publishedByKey = published.ToDictionary(entry => entry.Key, entry => entry.Row, StringComparer.OrdinalIgnoreCase);

            // A row whose identifying cells ALONE changed is one row renamed, not a removal and an
            // addition - BoardDataDiffer's own rule (owner decision, 2026-10-04).
            var renamedFrom = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (pairRenames)
            {
                List<(string Key, T Row)> removed = based.Where(entry => !draftByKey.ContainsKey(entry.Key)).ToList();
                List<(string Key, T Row)> added = draft.Where(entry => !baseByKey.ContainsKey(entry.Key)).ToList();

                foreach (BoardRowRename rename in BoardDataDiffer.PairRenamedRows(
                             removed.Select(entry => (object)entry.Row).ToList(),
                             added.Select(entry => (object)entry.Row).ToList()))
                {
                    renamedFrom[added[rename.Added].Key] = removed[rename.Removed].Key;
                }
            }

            var renamedAway = new HashSet<string>(renamedFrom.Values, StringComparer.OrdinalIgnoreCase);

            // Start from the published board, in its order.
            List<Slot<T>> result = published.Select(entry => new Slot<T>(entry.Key, entry.Row)).ToList();

            int IndexOf(string key) => result.FindIndex(slot => string.Equals(slot.Key, key, StringComparison.OrdinalIgnoreCase));

            string Describe(T row) => BoardDraftNaturalKeys.ForRow(row) is { Length: > 0 } key ? key : identity(row);

            void Conflict(T row, string what) =>
                state.Conflicts.Add(new DraftRebaseConflict(section, Describe(row), what));

            // 1. Rows the contributor removed (and did not rename): removed, whatever the update did.
            foreach ((string key, T row) in based)
            {
                if (draftByKey.ContainsKey(key) || renamedAway.Contains(key))
                    continue;

                state.ChangesKept++;

                if (publishedByKey.TryGetValue(key, out T? publishedRow) && !BoardDataDiffer.SameValues(publishedRow, row))
                    Conflict(row, "The update changed this row, and your draft removes it. It stays removed.");

                int at = IndexOf(key);

                if (at >= 0)
                    result.RemoveAt(at);
            }

            // 2. Rows in both the base and the draft: left alone, or changed by the contributor.
            // 3. Rows the contributor renamed.  4. Rows the contributor added.
            for (int position = 0; position < draft.Count; position++)
            {
                (string key, T row) = draft[position];

                if (baseByKey.TryGetValue(key, out T? basedRow))
                {
                    if (BoardDataDiffer.SameValues(row, basedRow))
                        continue; // Left alone: the published update's version (or removal) stands.

                    state.ChangesKept++;

                    int existingRow = IndexOf(key);

                    if (existingRow >= 0)
                    {
                        T publishedRow = result[existingRow].Row;

                        if (!BoardDataDiffer.SameValues(publishedRow, basedRow) && !BoardDataDiffer.SameValues(publishedRow, row))
                            Conflict(row, "The update and your draft both changed this row. Your version is kept.");

                        result[existingRow].Row = row;
                    }
                    else
                    {
                        Conflict(row, "The update removed this row, and your draft changes it. Your version is kept.");
                        DraftRebase.Insert(result, draft, position, key, row);
                    }

                    continue;
                }

                state.ChangesKept++;

                if (renamedFrom.TryGetValue(key, out string? oldKey))
                {
                    int existingNew = IndexOf(key);
                    int existingOld = IndexOf(oldKey);

                    if (existingNew >= 0)
                    {
                        // The update added a row under the very name the contributor renamed to.
                        Conflict(row, "The update added a row with the name your draft gives this row. Your version is kept.");
                        result[existingNew].Row = row;

                        if (existingOld >= 0)
                            result.RemoveAt(existingOld);
                    }
                    else if (existingOld >= 0)
                    {
                        if (!BoardDataDiffer.SameValues(result[existingOld].Row, baseByKey[oldKey]))
                            Conflict(row, "The update changed this row, and your draft renames it. Your version is kept.");

                        result[existingOld].Key = key;
                        result[existingOld].Row = row;
                    }
                    else
                    {
                        Conflict(row, "The update removed this row, and your draft renames it. Your version is kept.");
                        DraftRebase.Insert(result, draft, position, key, row);
                    }

                    continue;
                }

                int existing = IndexOf(key);

                if (existing >= 0)
                {
                    if (!BoardDataDiffer.SameValues(result[existing].Row, row))
                        Conflict(row, "The update added this row too, with other values. Your version is kept.");

                    result[existing].Row = row;
                }
                else
                {
                    DraftRebase.Insert(result, draft, position, key, row);
                }
            }

            return result.Select(slot => slot.Row).ToList();
        }

        // ###########################################################################################
        // The rows the published update RENAMED - identifying cells alone changed, BoardDataDiffer's
        // pairing from the base to the published board - given their new identity in the base and
        // in the draft, before the merge (see the class header). The base's row becomes the
        // published one (the same row, renamed); the contributor's keeps its own values under the
        // update's key. A new key the draft already holds (the contributor added a row of that name)
        // is not applied to the draft: the merge then lists it as the conflict it is.
        // ###########################################################################################
        private static void FollowUpdateRenames<T>(
            List<(string Key, T Row)> based,
            List<(string Key, T Row)> draft,
            List<(string Key, T Row)> published)
            where T : class
        {
            var publishedKeys = new HashSet<string>(published.Select(entry => entry.Key), StringComparer.OrdinalIgnoreCase);
            var baseKeys = new HashSet<string>(based.Select(entry => entry.Key), StringComparer.OrdinalIgnoreCase);

            List<int> removed = Enumerable.Range(0, based.Count).Where(index => !publishedKeys.Contains(based[index].Key)).ToList();
            List<int> added = Enumerable.Range(0, published.Count).Where(index => !baseKeys.Contains(published[index].Key)).ToList();

            if (removed.Count == 0 || added.Count == 0)
                return;

            var draftKeys = new HashSet<string>(draft.Select(entry => entry.Key), StringComparer.OrdinalIgnoreCase);

            foreach (BoardRowRename rename in BoardDataDiffer.PairRenamedRows(
                         removed.Select(index => (object)based[index].Row).ToList(),
                         added.Select(index => (object)published[index].Row).ToList()))
            {
                int baseIndex = removed[rename.Removed];
                (string oldKey, _) = based[baseIndex];
                (string newKey, T newRow) = published[added[rename.Added]];

                if (draftKeys.Contains(newKey))
                    continue;

                int draftIndex = draft.FindIndex(entry => string.Equals(entry.Key, oldKey, StringComparison.OrdinalIgnoreCase));

                if (draftIndex >= 0)
                {
                    if (DraftRebase.WithKeyOf(draft[draftIndex].Row, newRow) is not T relabelled)
                        continue; // A row type that cannot be copied: merged as before.

                    draft[draftIndex] = (newKey, relabelled);
                    draftKeys.Remove(oldKey);
                    draftKeys.Add(newKey);
                }

                based[baseIndex] = (newKey, newRow);
            }
        }

        // ###########################################################################################
        // `row` with the identifying cells (BoardDraftNaturalKeys) of `keySource` - a copy, the rows
        // being init-only. Null for a row type that cannot be copied (no parameterless constructor).
        // ###########################################################################################
        private static object? WithKeyOf(object row, object keySource)
        {
            Type type = row.GetType();

            if (type != keySource.GetType() || type.GetConstructor(Type.EmptyTypes) is null)
                return null;

            IReadOnlyList<string> keyProperties = BoardDraftNaturalKeys.PropertiesOf(type);
            object copy = Activator.CreateInstance(type)!;

            foreach (PropertyInfo property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!property.CanRead || !property.CanWrite || property.GetIndexParameters().Length > 0)
                    continue;

                object source = keyProperties.Contains(property.Name, StringComparer.Ordinal) ? keySource : row;
                property.SetValue(copy, property.GetValue(source));
            }

            return copy;
        }

        // ###########################################################################################
        // Each row with its identity: its key plus its occurrence among rows sharing that key
        // ("U8#0", "U8#1"), so duplicates pair in order and none is lost.
        // ###########################################################################################
        private static List<(string Key, T Row)> Keyed<T>(IReadOnlyList<T>? rows, Func<T, string> identity)
            where T : class
        {
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var keyed = new List<(string Key, T Row)>();

            foreach (T row in rows ?? [])
            {
                if (row is null)
                    continue;

                string key = identity(row) ?? string.Empty;
                int occurrence = seen.TryGetValue(key, out int count) ? count : 0;
                seen[key] = occurrence + 1;

                keyed.Add(($"{key}#{occurrence}", row));
            }

            return keyed;
        }

        // A row the contributor added goes after the row it followed in the draft (the nearest
        // earlier draft row still in the result), else before the nearest later one, else last.
        private static void Insert<T>(List<Slot<T>> result, List<(string Key, T Row)> draft, int position, string key, T row)
        {
            int IndexOf(string wanted) => result.FindIndex(slot => string.Equals(slot.Key, wanted, StringComparison.OrdinalIgnoreCase));

            for (int before = position - 1; before >= 0; before--)
            {
                int at = IndexOf(draft[before].Key);

                if (at >= 0)
                {
                    result.Insert(at + 1, new Slot<T>(key, row));
                    return;
                }
            }

            for (int after = position + 1; after < draft.Count; after++)
            {
                int at = IndexOf(draft[after].Key);

                if (at >= 0)
                {
                    result.Insert(at, new Slot<T>(key, row));
                    return;
                }
            }

            result.Add(new Slot<T>(key, row));
        }
    }
}
