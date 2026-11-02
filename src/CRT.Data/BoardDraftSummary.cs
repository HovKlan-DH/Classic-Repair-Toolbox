using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What the UI needs to know about a draft WITHOUT re-deriving it from BoardDraftApplier's
    // merge (NewContributeStrategy.md Phase 2, session 2b, task 5 - "mark every drafted row
    // visibly"). BoardDraftApplier.ApplyDraft deliberately throws away which merged row came from
    // the draft once it has produced the merged BoardData (see its own header comment) - the
    // merged result is meant to be indistinguishable from a board that was always this way. A
    // renderer that wants to mark a row as drafted needs a second, cheap answer to "is this row
    // drafted", not a change to what the merge itself produces.
    //
    // Built from the COMPARISON between the drafted board and the published one - see FromChanges,
    // which is the only way in since Phase 6 retired the delta model (2026-09-23). It is a couple
    // of HashSets, built once per board load alongside a comparison that has already been made.
    //
    // Deliberately narrow: only what has a UI marker - the component list's chip, for every row that is a
    // component's (code review, 2026-11-02), and the schematic tint, for its highlights. A future section
    // needing the same treatment adds its own set here rather than this class trying to generalize
    // ahead of a second real need - see BoardDraftApplier's own KeyOf<T> for the shape a fully
    // generic version would have to take, which is more machinery than two sets justify today.
    // ###########################################################################################
    public sealed class BoardDraftSummary
    {
        public static readonly BoardDraftSummary Empty = new(
            new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            new HashSet<string>(StringComparer.OrdinalIgnoreCase));

        // Board labels with a Modified or Added row in the draft's Components section, with any row of theirs
        // changed in the sections they hold beside it (ComponentRowSections), or with their own rows in
        // another order - what the component list checks per row to decide whether to show the "drafted" chip.
        public IReadOnlySet<string> DraftedComponentBoardLabels { get; }

        // "<schematic>␟<label>" keys (BoardDraftNaturalKeys.ForComponentHighlight's own
        // shape) with a Modified or Added row in the draft's ComponentHighlights section - what
        // the schematic overlay checks per highlight rect to decide whether to tint it.
        public IReadOnlySet<string> DraftedHighlightKeys { get; }

        public bool HasAnyMarkers => this.DraftedComponentBoardLabels.Count > 0 || this.DraftedHighlightKeys.Count > 0;

        private BoardDraftSummary(HashSet<string> draftedComponentBoardLabels, HashSet<string> draftedHighlightKeys)
        {
            this.DraftedComponentBoardLabels = draftedComponentBoardLabels;
            this.DraftedHighlightKeys = draftedHighlightKeys;
        }

        // ###########################################################################################
        // Builds the same summary from a COMPARISON instead of from a draft's recorded deltas
        // (NewContributeStrategy.md Phase 6 - drafts stored as real board folders, 2026-09-23).
        //
        // *** WHY THERE ARE NOW TWO WAYS IN. *** FromDraft below reads a ledger written at EDIT
        // time. That cannot survive the contributor editing the board workbook in Excel with the
        // application closed, which is exactly what the new draft layout allows - so the drafted
        // rows are DERIVED by comparing the draft workbook against the published one
        // (BoardDataDiffer), and this turns that comparison into the same two sets the UI has
        // always read.
        //
        // Nothing about the rendering changes: the component list still checks
        // DraftedComponentBoardLabels per row, the schematic overlay still checks
        // DraftedHighlightKeys. Only where the answer comes from has moved.
        //
        // *** DELETED ROWS ARE EXCLUDED, for the same reason FromDraft excludes them: *** a row
        // that is absent from the draft is not ON the rendered board, so there is nothing on
        // screen to mark. Marking it would mean drawing a chip for a component the contributor
        // has removed. That is its row on the Components sheet, and the rectangle a highlight tints.
        //
        // *** A COMPONENT'S OTHER ROWS MARK IT, TAKEN OFF IT TOO (code review, 2026-11-02). *** Only its row
        // on the Components sheet did, so a change of its pins, replacement parts, files, links, pictures,
        // baselines or rectangles took "Draft" off it as soon as Main showed the board again - and the
        // Schematics tab's list never had it - while "Review changes" listed it. A pin deleted is a change
        // of a component still there, and one that went altogether is in no list to mark.
        //
        // *** A ROW GIVEN TO ANOTHER COMPONENT MARKS BOTH (code review, 2026-11-02): *** a pin's Board label
        // changed from U1 to U2 in the table, U1 has its "Undo changes to" and a group "1 pin removed" too -
        // so a paired row marks the label it had (PreviousNaturalKey) as well as the one it has.
        //
        // *** BUT A RECTANGLE THAT ONLY FOLLOWED ITS SCHEMATIC IS THE BOARD'S CHANGE, as ComponentRevert has it
        // - the image renamed, or removed with what was marked on it. *** Every component marked there would
        // read "Draft" for it, with nothing of its own to take back. Still tinted while it is on screen. It
        // followed one only when the board has left that name: a schematic spelt in other capitals in the
        // rectangle's row alone, or a rectangle handed to a sheet still there, is a change of the component.
        //
        // `reorderedComponents` is what the comparison cannot say (BoardRowOrder.ComponentsReordered): a
        // component whose own rows only moved has no changed row, and says "Draft" all the same.
        // ###########################################################################################
        public static BoardDraftSummary FromChanges(IEnumerable<BoardRowChange>? changes, IReadOnlySet<string>? reorderedComponents = null)
        {
            if (changes is null)
            {
                return Empty;
            }

            List<BoardRowChange> rows = changes.Where(change => change is not null).ToList();

            // The names the board no longer has: a schematic removed, or renamed.
            var schematicNamesLeft = new HashSet<string>(
                rows.Where(change => string.Equals(change.Section, "Board schematics", StringComparison.Ordinal))
                    .Select(change => change.Kind == BoardRowChangeKind.Deleted ? change.NaturalKey : change.PreviousNaturalKey)
                    .Where(name => name.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            var draftedComponentBoardLabels = new HashSet<string>(reorderedComponents ?? new HashSet<string>(), StringComparer.OrdinalIgnoreCase);
            var draftedHighlightKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string PartOf(string? naturalKey, int part)
            {
                string[] parts = (naturalKey ?? string.Empty).Split(BoardDraftNaturalKeys.Separator);

                return part < parts.Length ? parts[part].Trim() : string.Empty;
            }

            void Mark(string? naturalKey, int labelPart)
            {
                string label = PartOf(naturalKey, labelPart);

                if (label.Length > 0)
                {
                    draftedComponentBoardLabels.Add(label);
                }
            }

            foreach (BoardRowChange change in rows)
            {

                // Matched on the SECTION NAME the differ emits, which is DraftDriftReport's own
                // vocabulary - the two were deliberately kept identical so a section cannot be
                // called one thing in the change list and another here.
                if (string.Equals(change.Section, "Components", StringComparison.Ordinal))
                {
                    // The board label is the key's FIRST part. It used to be the whole key, but a
                    // regionalised component's key is "label U+241F region" since 2026-09-24 (see
                    // BoardDraftNaturalKeys.ForComponent), and the chip marks the component, not
                    // one of its regions.
                    if (change.Kind != BoardRowChangeKind.Deleted)
                    {
                        Mark(change.NaturalKey, labelPart: 0);
                        Mark(change.PreviousNaturalKey, labelPart: 0);
                    }
                }
                else if (BoardDraftSummary.ComponentRowSections.Contains(change.Section))
                {
                    // The label first in each of their keys, as on the Components sheet.
                    Mark(change.NaturalKey, labelPart: 0);
                    Mark(change.PreviousNaturalKey, labelPart: 0);
                }
                else if (string.Equals(change.Section, "Component highlights", StringComparison.Ordinal))
                {
                    string schematicBefore = PartOf(change.PreviousNaturalKey.Length > 0 ? change.PreviousNaturalKey : change.NaturalKey, 0);

                    bool followedItsSchematic = schematicNamesLeft.Contains(schematicBefore) &&
                        (change.Kind == BoardRowChangeKind.Deleted ||
                         (change.ChangedFields.Count > 0 && change.ChangedFields.All(field => field == nameof(ComponentHighlightEntry.SchematicName))));

                    // "<schematic>U+241F<label>": the label second.
                    if (!followedItsSchematic)
                    {
                        Mark(change.NaturalKey, labelPart: 1);
                        Mark(change.PreviousNaturalKey, labelPart: 1);
                    }

                    // Already the shape the overlay looks up - BoardDraftNaturalKeys built it. Carried
                    // through verbatim rather than rebuilt.
                    if (change.Kind != BoardRowChangeKind.Deleted && !string.IsNullOrWhiteSpace(change.NaturalKey))
                    {
                        draftedHighlightKeys.Add(change.NaturalKey);
                    }
                }
            }

            if (draftedComponentBoardLabels.Count == 0 && draftedHighlightKeys.Count == 0)
            {
                return Empty;
            }

            return new BoardDraftSummary(draftedComponentBoardLabels, draftedHighlightKeys);
        }

        // The sections a component holds beside its row, whose keys start with its label (BoardDraftNaturalKeys) -
        // the names BoardDataDiffer.Compare gives them. Never Credits or the board's files and links: their
        // keys start with something else, and would go into a set the component list searches by label.
        private static readonly HashSet<string> ComponentRowSections = new(StringComparer.Ordinal)
        {
            "Component images",
            "Component local files",
            "Component links",
            BoardWorkbookSchema.SheetComponentPins,
            BoardWorkbookSchema.SheetComponentReplacements,
        };
    }
}
