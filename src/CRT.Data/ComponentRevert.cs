using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE COMPONENT BACK TO ITS PUBLISHED VALUES (CRT 4.0 step 1, 2026-10-25 - the plan: "Each way of
    // throwing work away says what it covers: Undo - one step. Undo changes to IC1 - one component,
    // back to its published values. Discard draft - everything").
    //
    // Every row of the component - its rows in every region, its pictures and scope baselines, its
    // marks on the schematics, its files and links, its pins and replacement parts - is replaced by
    // the published board's rows of
    // it; a component the published board does not have goes altogether. A component the draft
    // RELABELLED (ComponentRelabel, the rule its highlights follow) goes back under its old label,
    // so "Undo changes to U08" brings U8 back. The rest of the draft is left exactly as it is.
    //
    // The rows go back WHERE THEY WERE: in place of the draft's rows of the component, or - when the
    // draft removed it - after the row the published board has before them. Order is no change the
    // differ counts, but a component put back at the end of its sheet would read as moved.
    //
    // A FILE OF THE COMPONENT REPLACED UNDER ITS OWN NAME goes back too (code review, 2026-10-26) -
    // not by the rows, which name it the same, but by the draft's list of files, which the caller
    // puts back after the rows (ReplacedFiles names them; DraftEdits.PutBackPublishedFilesAsync).
    // That part is a write of the file list, so the draft's Undo takes back the rows and not it.
    //
    // Pure: two boards in, a board out. One undoable step in the draft (DraftEdits).
    // ###########################################################################################
    public static class ComponentRevert
    {
        // ###########################################################################################
        // Whether `label` differs from the published board at all - the "Undo changes to" button is
        // only offered then. With the two board files' lists (`publishedFiles`, `draftFiles`), a file
        // of the component replaced under its own name is a change too (ReplacedFiles).
        // ###########################################################################################
        public static bool HasChanges(
            BoardData draft,
            BoardData? published,
            string? label,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? publishedFiles = null,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? draftFiles = null)
        {
            ArgumentNullException.ThrowIfNull(draft);

            BoardData reverted = ComponentRevert.ToPublished(draft, published, label);

            return BoardDataDiffer.Compare(draft, reverted).Count > 0 ||
                   ComponentRevert.ReplacedFiles(draft, published, label, publishedFiles, draftFiles).Count > 0;
        }

        // ###########################################################################################
        // `draft` with the component's rows as published (ToPublished) - or null when they already
        // read as published: nothing to undo in them, and no empty step in the draft's history. A
        // file replaced under its own name may still be (ReplacedFiles).
        // ###########################################################################################
        public static BoardData? RowsToPublished(BoardData draft, BoardData? published, string? label)
        {
            ArgumentNullException.ThrowIfNull(draft);

            BoardData reverted = ComponentRevert.ToPublished(draft, published, label);

            return ReferenceEquals(reverted, draft) || BoardDataDiffer.Compare(draft, reverted).Count == 0 ? null : reverted;
        }

        // ###########################################################################################
        // THE COMPONENT'S FILES THE DRAFT REPLACED UNDER THEIR OWN NAMES (code review, 2026-10-26):
        // cited by its rows as published - a relabelled component's, its old label's - and listed by
        // the draft (`draftFiles`) with other bytes than the published board lists (`publishedFiles`).
        // The rows alone are then no change, so "Undo changes to" was not offered - and when it was,
        // it put the rows back and left the replaced bytes in the draft and its next submission. Put
        // back by the draft's file list (DraftWorkbookStore.RemoveFiles: a name the published board
        // has goes back to the published one). None without both lists or a published board.
        // ###########################################################################################
        public static List<string> ReplacedFiles(
            BoardData draft,
            BoardData? published,
            string? label,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? publishedFiles,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? draftFiles)
        {
            ArgumentNullException.ThrowIfNull(draft);

            string wanted = label?.Trim() ?? string.Empty;

            if (wanted.Length == 0 || published is null || publishedFiles is null || draftFiles is null)
                return [];

            HashSet<string> labels = ComponentRevert.LabelsOf(draft, published, wanted);
            bool Of(string? rowLabel) => labels.Contains((rowLabel ?? string.Empty).Trim());

            return published.ComponentImages.Where(row => Of(row.BoardLabel)).Select(row => row.File)
                .Concat(published.ComponentLocalFiles.Where(row => Of(row.BoardLabel)).Select(row => row.File))
                .Select(BoardFileNames.Normalise)
                .Where(name => name.Length > 0 &&
                               publishedFiles.TryGetValue(name, out BoardDocument.StoredFileRow? theirs) &&
                               draftFiles.TryGetValue(name, out BoardDocument.StoredFileRow? mine) &&
                               !string.Equals(theirs.Sha256, mine.Sha256, StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Order(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // The label and every label the draft relabelled into it (ComponentRelabel).
        private static HashSet<string> LabelsOf(BoardData draft, BoardData published, string wanted)
        {
            var labels = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { wanted };

            foreach ((string from, string to) in ComponentRelabel.Between(published.Components, draft.Components))
            {
                if (string.Equals(to, wanted, StringComparison.OrdinalIgnoreCase))
                    labels.Add(from);
            }

            return labels;
        }

        // ###########################################################################################
        // `draft` with the component `label` as the published board has it. The board unchanged
        // (the same object) for a blank label.
        // ###########################################################################################
        public static BoardData ToPublished(BoardData draft, BoardData? published, string? label)
        {
            ArgumentNullException.ThrowIfNull(draft);

            string wanted = label?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
                return draft;

            BoardData before = published ?? new BoardData { HardwareName = string.Empty, BoardName = string.Empty };

            HashSet<string> labels = ComponentRevert.LabelsOf(draft, before, wanted);

            bool Of(string? rowLabel) => labels.Contains((rowLabel ?? string.Empty).Trim());

            // ###########################################################################################
            // *** THE PUBLISHED HIGHLIGHTS UNDER THE DRAFT'S SCHEMATIC NAMES (code review, 2026-10-26). ***
            // A schematic the draft renamed (the same image file under a new name - the rule every draft
            // edit follows, SubmissionRowsBoard.FollowSchematics) keeps its new name: put back under the
            // published one, the highlights named a schematic the draft no longer has - refused at Submit
            // and fixable nowhere in the table. One on an image the draft removed goes with it.
            // ###########################################################################################
            IReadOnlyDictionary<string, string?> renamed = SubmissionRowsBoard.FollowSchematics(before.Schematics, draft.Schematics);
            BoardData publishedHere = renamed.Count == 0
                ? before
                : before.WithComponentHighlights(SchematicRename.FollowHighlights(before.ComponentHighlights, renamed));

            return new BoardData
            {
                RevisionDate = draft.RevisionDate,
                HardwareName = draft.HardwareName,
                BoardName = draft.BoardName,
                Schematics = draft.Schematics,
                Components = ComponentRevert.Restore(draft.Components, before.Components, row => Of(row.BoardLabel)),
                ComponentImages = ComponentRevert.Restore(draft.ComponentImages, before.ComponentImages, row => Of(row.BoardLabel)),
                ComponentHighlights = ComponentRevert.Restore(draft.ComponentHighlights, publishedHere.ComponentHighlights, row => Of(row.BoardLabel)),
                ComponentLocalFiles = ComponentRevert.Restore(draft.ComponentLocalFiles, before.ComponentLocalFiles, row => Of(row.BoardLabel)),
                ComponentLinks = ComponentRevert.Restore(draft.ComponentLinks, before.ComponentLinks, row => Of(row.BoardLabel)),
                ComponentPins = ComponentRevert.Restore(draft.ComponentPins, before.ComponentPins, row => Of(row.BoardLabel)),
                ComponentReplacements = ComponentRevert.Restore(draft.ComponentReplacements, before.ComponentReplacements, row => Of(row.BoardLabel)),
                BoardLocalFiles = draft.BoardLocalFiles,
                BoardLinks = draft.BoardLinks,
                Credits = draft.Credits,
                KiCadImportantSignals = draft.KiCadImportantSignals,
            };
        }

        // ###########################################################################################
        // One section: the draft's rows without the component's, the published rows of it put back
        // where the draft's first one was - or, with none, right after the row the published section
        // has before them (by identity, BoardDraftNaturalKeys), else first.
        // ###########################################################################################
        private static List<T> Restore<T>(List<T> draftRows, List<T> publishedRows, Func<T, bool> ofComponent)
            where T : class
        {
            List<T> restored = publishedRows.Where(ofComponent).ToList();
            int firstOwn = draftRows.FindIndex(row => ofComponent(row));

            List<T> kept = draftRows.Where(row => !ofComponent(row)).ToList();

            if (restored.Count == 0)
                return kept;

            int at;

            if (firstOwn >= 0)
            {
                // The draft's rows before the first of the component's that are kept.
                at = draftRows.Take(firstOwn).Count(row => !ofComponent(row));
            }
            else
            {
                int publishedFirst = publishedRows.FindIndex(row => ofComponent(row));

                var keptKeys = kept
                    .Select((row, index) => (Key: BoardDraftNaturalKeys.ForRow(row), Index: index))
                    .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First().Index, StringComparer.OrdinalIgnoreCase);

                at = 0;

                for (int index = publishedFirst - 1; index >= 0; index--)
                {
                    if (keptKeys.TryGetValue(BoardDraftNaturalKeys.ForRow(publishedRows[index]), out int keptIndex))
                    {
                        at = keptIndex + 1;
                        break;
                    }
                }
            }

            kept.InsertRange(at, restored);
            return kept;
        }
    }
}
