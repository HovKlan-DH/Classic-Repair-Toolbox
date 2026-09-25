using System;
using System.Collections.Generic;

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
    // Deliberately narrow: only the two sections that currently have a UI marker (Components for
    // the component list chip, ComponentHighlights for the schematic tint). A future section
    // needing the same treatment adds its own set here rather than this class trying to generalize
    // ahead of a second real need - see BoardDraftApplier's own KeyOf<T> for the shape a fully
    // generic version would have to take, which is more machinery than two sets justify today.
    // ###########################################################################################
    public sealed class BoardDraftSummary
    {
        public static readonly BoardDraftSummary Empty = new(
            new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            new HashSet<string>(StringComparer.OrdinalIgnoreCase));

        // Board labels with a Modified or Added row in the draft's Components section - what the
        // component list checks per row to decide whether to show the "drafted" chip.
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
        // has removed.
        // ###########################################################################################
        public static BoardDraftSummary FromChanges(IEnumerable<BoardRowChange>? changes)
        {
            if (changes is null)
            {
                return Empty;
            }

            var draftedComponentBoardLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var draftedHighlightKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardRowChange change in changes)
            {
                if (change is null || change.Kind == BoardRowChangeKind.Deleted)
                {
                    continue;
                }

                // Matched on the SECTION NAME the differ emits, which is DraftDriftReport's own
                // vocabulary - the two were deliberately kept identical so a section cannot be
                // called one thing in the change list and another here.
                if (string.Equals(change.Section, "Components", StringComparison.Ordinal))
                {
                    // The board label is the key's FIRST part. It used to be the whole key, but a
                    // regionalised component's key is "label U+241F region" since 2026-09-24 (see
                    // BoardDraftNaturalKeys.ForComponent), and the chip marks the component, not
                    // one of its regions.
                    string label = (change.NaturalKey ?? string.Empty)
                        .Split(BoardDraftNaturalKeys.Separator)[0]
                        .Trim();

                    if (label.Length > 0)
                    {
                        draftedComponentBoardLabels.Add(label);
                    }
                }
                else if (string.Equals(change.Section, "Component highlights", StringComparison.Ordinal))
                {
                    // Already "<schematic>U+241F<label>" - BoardDraftNaturalKeys built it, and it is
                    // the shape the overlay looks up. Carried through verbatim rather than rebuilt.
                    if (!string.IsNullOrWhiteSpace(change.NaturalKey))
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

    }
}
