using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE A COMPONENT'S ROWS GO IN A BOARD WORKBOOK (owner request, 2026-09-24).
    //
    // *** ROW ORDER IS WHAT THE USER SEES. *** The main window's component list, its category list
    // and the Overview tab all show components in the Components sheet's own order and never
    // re-sort (see ComponentListBuilder: categories are listed "in first-seen order"). So where a
    // writer puts a row is exactly where the component appears in the application.
    //
    // THE RULES, every writer going through them:
    //
    //   - A NEW component goes into its CATEGORY's group - the rows already carrying that category
    //     - in natural board-label order (C1, C2, C10, not C1, C10, C2; see NaturalLabelComparer).
    //     A category nobody uses yet goes at the END, becoming the last category in the list.
    //
    //   - An EXISTING component stays where it is when it is edited. It used to be removed and
    //     re-added at the bottom on every save from the Contribute window, so simply editing C1
    //     after adding C2 put C2 above C1 - the order a contributor saw drifted with every edit,
    //     which is what the project owner reported.
    //
    // Only NEW rows are placed. Nothing here ever re-sorts rows that are already in the file: a
    // workbook's existing order is the project owner's (or a contributor's, dragged into place in
    // the table editor), and rearranging it behind their back would be worse than the drift this
    // fixes.
    //
    // PURE: lists in, lists out.
    // ###########################################################################################
    public static class ComponentPlacement
    {
        // ###########################################################################################
        // The index a NEW component row should be inserted at:
        //
        //   - straight after the last row with the SAME board label, when there is one - a new
        //     regional variant (U1 for NTSC beside U1 for PAL) belongs with its twin, whatever
        //     category it was typed with;
        //   - otherwise before the first row of its category whose label sorts after it, else just
        //     after its category's last row, else at the end.
        //
        // *** THE SAME-LABEL RULE IS A REPORTED FIX (2026-09-24). *** A regional variant added in
        // the table editor with no category typed in "disappeared" on save: a blank category is a
        // category nobody used, so the row went to the very END of a long sheet, far from the
        // component it was a variant of.
        //
        // Labels and categories compare trimmed and case-insensitively, as everywhere else.
        // ###########################################################################################
        public static int InsertionIndex(IReadOnlyList<ComponentEntry> components, string? category, string? boardLabel)
        {
            ArgumentNullException.ThrowIfNull(components);

            string wantedLabel = boardLabel?.Trim() ?? string.Empty;

            if (wantedLabel.Length > 0)
            {
                for (int i = components.Count - 1; i >= 0; i--)
                {
                    if (string.Equals(components[i].BoardLabel?.Trim(), wantedLabel, StringComparison.OrdinalIgnoreCase))
                    {
                        return i + 1;
                    }
                }
            }

            string wantedCategory = category?.Trim() ?? string.Empty;
            int lastInCategory = -1;

            for (int i = 0; i < components.Count; i++)
            {
                if (!string.Equals(components[i].Category?.Trim() ?? string.Empty, wantedCategory, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (NaturalLabelComparer.Compare(components[i].BoardLabel, boardLabel) > 0)
                {
                    return i;
                }

                lastInCategory = i;
            }

            return lastInCategory >= 0 ? lastInCategory + 1 : components.Count;
        }

        // ###########################################################################################
        // The components with one component's rows replaced: IN PLACE when it already had rows,
        // otherwise at InsertionIndex. `replacements` stay together and in their own order (a
        // component can have one row per region).
        // ###########################################################################################
        public static List<ComponentEntry> ReplaceComponent(
            IReadOnlyList<ComponentEntry> components,
            string boardLabel,
            IReadOnlyList<ComponentEntry> replacements)
        {
            ArgumentNullException.ThrowIfNull(components);
            ArgumentNullException.ThrowIfNull(replacements);

            string label = boardLabel?.Trim() ?? string.Empty;

            return ComponentPlacement.ReplaceByLabel(
                components,
                label,
                entry => entry.BoardLabel,
                replacements,
                kept => replacements.Count == 0
                    ? kept.Count
                    : ComponentPlacement.InsertionIndex(kept, replacements[0].Category, replacements[0].BoardLabel));
        }

        // ###########################################################################################
        // The same for a component-scoped sheet with no category (images, local files, links): in
        // place when the component already has rows there, otherwise before the first row whose
        // label sorts after it - so a new component's files sit among their neighbours rather than
        // collecting at the bottom.
        // ###########################################################################################
        public static List<T> ReplaceRowsForLabel<T>(
            IReadOnlyList<T> rows,
            string boardLabel,
            Func<T, string> labelOf,
            IReadOnlyList<T> replacements)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(labelOf);
            ArgumentNullException.ThrowIfNull(replacements);

            string label = boardLabel?.Trim() ?? string.Empty;

            return ComponentPlacement.ReplaceByLabel(
                rows,
                label,
                labelOf,
                replacements,
                kept => ComponentPlacement.IndexByLabel(kept, label, labelOf));
        }

        // Before the first row whose label sorts after `label`, else at the end.
        public static int IndexByLabel<T>(IReadOnlyList<T> rows, string? label, Func<T, string> labelOf)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(labelOf);

            for (int i = 0; i < rows.Count; i++)
            {
                if (NaturalLabelComparer.Compare(labelOf(rows[i]), label) > 0)
                {
                    return i;
                }
            }

            return rows.Count;
        }

        // ###########################################################################################
        // Removes every row of `label` (case-insensitive) and puts `replacements` where the first of
        // them was - or, when there were none, where `placeNew` says. A BLANK label removes nothing:
        // it identifies no component, and treating it as a wildcard would clear every row whose
        // own label happened to be blank too.
        // ###########################################################################################
        private static List<T> ReplaceByLabel<T>(
            IReadOnlyList<T> rows,
            string label,
            Func<T, string> labelOf,
            IReadOnlyList<T> replacements,
            Func<List<T>, int> placeNew)
        {
            var kept = new List<T>(rows.Count + replacements.Count);
            int firstRemoved = -1;

            foreach (T row in rows)
            {
                bool isThisComponent = label.Length > 0 && string.Equals(
                    labelOf(row)?.Trim(),
                    label,
                    StringComparison.OrdinalIgnoreCase);

                if (isThisComponent)
                {
                    if (firstRemoved < 0)
                    {
                        firstRemoved = kept.Count;
                    }

                    continue;
                }

                kept.Add(row);
            }

            int insertAt = firstRemoved >= 0 ? firstRemoved : placeNew(kept);
            kept.InsertRange(Math.Clamp(insertAt, 0, kept.Count), replacements);

            return kept;
        }
    }
}
