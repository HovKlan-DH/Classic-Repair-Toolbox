using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The user's chosen top-to-bottom order for the schematic previews in the Workbooks tab's board
    // pane, and the maths for dragging one of them to a new position.
    //
    // WHY AN ORDER IS STORED AT ALL: the pane was sorted alphabetically by schematic name, which is
    // an order nobody chose and which rarely matches how a repair is actually worked - the board
    // someone is on now belongs at the top, not wherever its name happens to fall. So the previews
    // are draggable, exactly as an entry's photos are in the full editor, and the result is
    // persisted per workbook (WorkbookRecord.SchematicOrder).
    //
    // PER WORKBOOK, not per board: the pane only ever shows one workbook's schematics, and two jobs
    // on the same board are usually about different parts of it. The order travels in the workbook's
    // own index.json, so it is deleted along with the workbook and needs no cleanup of its own.
    //
    // NAMES, NOT INDICES: the stored list holds schematic NAMES. Which schematics appear in the pane
    // changes as worklogs are added and deleted, and a search filters it further, so a positional
    // list would silently re-order the pane the moment its length stopped matching. A name that no
    // longer has any entries simply does not appear, and its place in the stored list is harmless -
    // it costs one dead string and correctly restores the position if that schematic gets a worklog
    // again later.
    //
    // Comparison is case-insensitive throughout, matching how every other schematic-name lookup in
    // this app works (see TabWorkbooks.BoardPreviews.cs's own grouping) - board Excel files arrive
    // from the server independently of app releases and nothing normalises their casing.
    //
    // Pure list logic, no Avalonia and no disk: the tab reads it, the tests drive it directly.
    // ###########################################################################################
    public static class WorkbookSchematicOrder
    {
        // ###########################################################################################
        // Puts the schematic names actually being shown into the user's stored order.
        //
        // Names present in storedOrder come first, in that order. Anything else - a schematic whose
        // first worklog was just added, or every schematic when nothing has ever been dragged -
        // follows in the caller's own order, which is the alphabetical grouping the pane arrives
        // with. That "known first, then the rest" rule is what makes an empty stored order a no-op
        // rather than a special case, so a workbook nobody has dragged in behaves exactly as it did
        // before this existed.
        //
        // A NEW schematic sorts to the BOTTOM rather than into its alphabetical place among ordered
        // ones. Slotting it into the middle would push everything below it down by one for a
        // schematic the user has never positioned; arriving at the end is both predictable and one
        // drag away from wherever it belongs.
        //
        // Duplicate names in storedOrder are ignored after their first occurrence, and a stored name
        // that is not being shown is skipped - neither can be produced by ApplyMove, but index.json
        // is a plain file a user can hand-edit, so neither may corrupt the display either.
        // ###########################################################################################
        public static List<string> Apply(IEnumerable<string> shownNames, IEnumerable<string>? storedOrder)
        {
            var shown = shownNames?.ToList() ?? new List<string>();
            if (shown.Count == 0)
            {
                return new List<string>();
            }

            if (storedOrder == null)
            {
                return shown;
            }

            var remaining = new List<string>(shown);
            var ordered = new List<string>(shown.Count);

            foreach (var name in storedOrder)
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                int index = remaining.FindIndex(candidate =>
                    string.Equals(candidate, name, StringComparison.OrdinalIgnoreCase));

                if (index < 0)
                {
                    continue;
                }

                ordered.Add(remaining[index]);
                remaining.RemoveAt(index);
            }

            ordered.AddRange(remaining);
            return ordered;
        }

        // ###########################################################################################
        // Moves one schematic to a new position and returns the full order to store.
        //
        // Takes the names CURRENTLY SHOWN, in the order they are currently displayed, so the result
        // is exactly what the user sees after the drop - the caller does not have to reason about
        // how the previous stored order and the current set combine (Apply has already done that).
        //
        // The returned list is the complete new stored order and REPLACES the old one. Names from
        // the previous stored order that are not currently shown are deliberately dropped: keeping
        // them would mean deciding where each unshown name sits relative to a list the user has just
        // rearranged, and there is no answer to that which is not a guess. Losing the remembered
        // position of a schematic with no worklogs is a far smaller cost than silently reshuffling
        // the pane the next time one comes back.
        //
        // targetIndex is clamped, and a move to where the schematic already is returns the order
        // unchanged - so the caller can persist unconditionally without first testing for a no-op.
        // ###########################################################################################
        public static List<string> ApplyMove(IEnumerable<string> shownNamesInDisplayOrder, string movedName, int targetIndex)
        {
            var ordered = shownNamesInDisplayOrder?.ToList() ?? new List<string>();

            if (ordered.Count < 2 || string.IsNullOrWhiteSpace(movedName))
            {
                return ordered;
            }

            int currentIndex = ordered.FindIndex(candidate =>
                string.Equals(candidate, movedName, StringComparison.OrdinalIgnoreCase));

            if (currentIndex < 0)
            {
                return ordered;
            }

            int clampedTarget = Math.Clamp(targetIndex, 0, ordered.Count - 1);
            if (clampedTarget == currentIndex)
            {
                return ordered;
            }

            string moved = ordered[currentIndex];
            ordered.RemoveAt(currentIndex);
            ordered.Insert(clampedTarget, moved);

            return ordered;
        }

        // ###########################################################################################
        // Which slot a drop at a given Y lands in, over previews whose heights DIFFER (a schematic
        // preview is as tall as its image needs, unlike the uniform rows of the editor's photo list -
        // which is why this cannot be the simple "pointerY / rowHeight" that list can use).
        //
        // Walks the heights and returns the index of the first preview whose vertical MIDPOINT the
        // pointer has not yet passed - the standard "insert before or after this one" rule, which is
        // what makes a drop land where the gap is being shown rather than where the pointer happens
        // to be inside a tall image.
        //
        // Heights are the arranged heights of the previews in their current display order, INCLUDING
        // the one being dragged, and the result indexes into that same list. spacing is the gap
        // between previews, which is part of the distance the pointer has to travel and so must be
        // counted or every midpoint below the first drifts upward.
        // ###########################################################################################
        public static int ResolveDropIndex(IReadOnlyList<double> previewHeights, double spacing, double pointerY)
        {
            if (previewHeights == null || previewHeights.Count == 0)
            {
                return 0;
            }

            double top = 0;

            for (int i = 0; i < previewHeights.Count; i++)
            {
                double height = previewHeights[i];

                if (pointerY < top + (height / 2.0))
                {
                    return i;
                }

                top += height + spacing;
            }

            return previewHeights.Count - 1;
        }
    }
}
