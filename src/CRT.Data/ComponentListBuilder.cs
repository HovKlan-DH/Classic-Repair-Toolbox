using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Lightweight view model for a component list item.
    //
    // IsDrafted (NewContributeStrategy.md Phase 2, session 2b) marks a row whose Components entry
    // was added or modified by a local draft rather than being officially published - the main
    // window's component list renders it with a small chip. Decided by the CALLER
    // (BuildComponentItems below takes the drafted-labels set as a parameter) rather than by this
    // class reaching for BoardDraftSummary itself, so ComponentListBuilder stays usable with no
    // draft in play at all - most boards have none.
    // ###########################################################################################
    public sealed class ComponentListItem
    {
        public string DisplayText { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string SelectionKey { get; init; } = string.Empty;
        public bool IsDrafted { get; init; }
        public override string ToString() => this.DisplayText;
    }

    // ###########################################################################################
    // One row of the worklog entry card's "Mark components in scope" checklist: a component's
    // board label kept apart from its friendly/technical name, since the UI bolds only the label.
    // ###########################################################################################
    public sealed class ComponentInScope
    {
        public string BoardLabel { get; init; } = string.Empty;
        public string DisplayName { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // Builds the main window's component list: the region filter, the category filter, the
    // multi-term search, and the composed display text and selection key for each row.
    //
    // Extracted from Main so the filtering rules can be tested. Two rules matter and are easy to
    // break by accident: a component with a blank region is shared and shows in every region,
    // and search terms are ANDed across the whole composed display string rather than matched
    // against any single field.
    // ###########################################################################################
    public static class ComponentListBuilder
    {
        // ###########################################################################################
        // Returns the distinct component categories in first-seen order, ignoring blank ones.
        // ###########################################################################################
        public static List<string> BuildDistinctCategories(BoardData boardData)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var categories = new List<string>();

            foreach (var component in boardData.Components)
            {
                if (!string.IsNullOrWhiteSpace(component.Category) && seen.Add(component.Category))
                    categories.Add(component.Category);
            }

            return categories;
        }

        // ###########################################################################################
        // Builds component list items filtered by the given region and search string.
        //
        // draftedBoardLabels (session 2b) marks which rows to flag as drafted - typically
        // DataManager.LastLoadedDraftSummary.DraftedComponentBoardLabels, resolved by the caller
        // rather than here so this stays testable with no DataManager/DraftManager in play. Null
        // is treated the same as "no board labels are drafted", the ordinary case for a board with
        // no local draft.
        // ###########################################################################################
        public static List<ComponentListItem> BuildComponentItems(
            BoardData boardData,
            string region,
            HashSet<string>? categoryFilter = null,
            string searchTerm = "",
            IReadOnlySet<string>? draftedBoardLabels = null)
        {
            var items = new List<ComponentListItem>();

            var searchTerms = string.IsNullOrWhiteSpace(searchTerm)
                ? Array.Empty<string>()
                : searchTerm.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var component in boardData.Components)
            {
                var componentRegion = component.Region?.Trim() ?? string.Empty;

                if (!string.IsNullOrEmpty(componentRegion) &&
                    !string.Equals(componentRegion, region, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (categoryFilter != null && !categoryFilter.Contains(component.Category ?? string.Empty))
                    continue;

                var parts = new List<string>(3);
                if (!string.IsNullOrWhiteSpace(component.BoardLabel))
                    parts.Add(component.BoardLabel.Trim());
                if (!string.IsNullOrWhiteSpace(component.FriendlyName))
                    parts.Add(component.FriendlyName.Trim());
                if (!string.IsNullOrWhiteSpace(component.TechnicalNameOrValue))
                    parts.Add(component.TechnicalNameOrValue.Trim());

                if (parts.Count == 0)
                    continue;

                string displayString = string.Join(" | ", parts);

                if (searchTerms.Length > 0)
                {
                    bool matches = true;
                    foreach (var term in searchTerms)
                    {
                        if (displayString.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            matches = false;
                            break;
                        }
                    }

                    if (!matches)
                        continue;
                }

                string boardLabel = component.BoardLabel?.Trim() ?? string.Empty;

                items.Add(new ComponentListItem
                {
                    BoardLabel = boardLabel,
                    DisplayText = displayString,
                    SelectionKey = string.Join("\u001F",
                        boardLabel,
                        component.FriendlyName?.Trim() ?? string.Empty,
                        component.TechnicalNameOrValue?.Trim() ?? string.Empty,
                        component.Region?.Trim() ?? string.Empty),
                    IsDrafted = boardLabel.Length > 0 && (draftedBoardLabels?.Contains(boardLabel) ?? false)
                });
            }

            return items;
        }

        // ###########################################################################################
        // Builds the "Mark components in scope" checklist rows: every component whose board label is
        // in boardLabelsInScope, kept in boardData.Components order (the same order Overview and the
        // main Component list use - neither re-sorts) rather than the order boardLabelsInScope was
        // built in.
        //
        // Each board label yields at most ONE row. A regionalized component occupies several rows in
        // the board Excel - one per region, all sharing a label - and listing it once per row showed
        // the same physical part two or three times with nothing to tell the duplicates apart. The
        // checklist marks physical components, and there is only one U1 on the board.
        //
        // The first row seen supplies the display name, the same way BuildDistinctCategories keeps
        // the first spelling of a category. No row is more correct than the others here: the scope
        // is decided by highlight rectangles, which are keyed by board label with no region of their
        // own, so this list has no region to filter by in the first place. (BuildComponentItems
        // above CAN filter, because the main list is explicitly showing one region at a time.)
        // ###########################################################################################
        public static List<ComponentInScope> BuildComponentsInScope(BoardData boardData, IReadOnlyCollection<string> boardLabelsInScope)
        {
            // Always rebuilt with OrdinalIgnoreCase rather than trusting the caller's set, so
            // matching stays case-insensitive - board label comparisons are case-insensitive
            // everywhere else in the app (e.g. the schematic hover lookups) - regardless of which
            // comparer the caller's HashSet happened to be constructed with.
            var scope = new HashSet<string>(boardLabelsInScope, StringComparer.OrdinalIgnoreCase);
            var results = new List<ComponentInScope>();

            // Case-insensitive to match `scope` and every other board-label comparison in the app,
            // so rows spelled "U1" and "u1" are recognised as the one component they describe.
            var alreadyListed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var component in boardData.Components)
            {
                string boardLabel = component.BoardLabel?.Trim() ?? string.Empty;

                // A blank label identifies nothing, so it can neither be matched against the scope
                // nor de-duplicated meaningfully - several unrelated components would collapse into
                // one row showing an empty label. Skipped outright rather than listed.
                if (boardLabel.Length == 0)
                    continue;

                if (!scope.Contains(boardLabel))
                    continue;

                if (!alreadyListed.Add(boardLabel))
                    continue;

                var parts = new List<string>(2);
                if (!string.IsNullOrWhiteSpace(component.FriendlyName))
                    parts.Add(component.FriendlyName.Trim());
                if (!string.IsNullOrWhiteSpace(component.TechnicalNameOrValue))
                    parts.Add(component.TechnicalNameOrValue.Trim());

                results.Add(new ComponentInScope
                {
                    BoardLabel = boardLabel,
                    DisplayName = string.Join(" | ", parts)
                });
            }

            return results;
        }

        // ###########################################################################################
        // Narrows a worklog entry's selected components to those its marked area still touches.
        //
        // Called when an area is resized on the schematic. The rule is deliberately one-directional:
        //
        //   - a label the area NO LONGER touches is dropped, because the user's selection said
        //     "this component is in scope" about an area that no longer covers it, and leaving it
        //     would silently keep a component associated with a fault it is no longer part of;
        //   - a label the area NOW touches is NOT added, because being inside the rectangle is not
        //     the same as the user having decided it is relevant. Auto-ticking would quietly put
        //     components into someone's worklog that they never chose, and the wider the area is
        //     dragged the more of them appear.
        //
        // So a resize can only ever remove. Adding stays a deliberate act in the full editor's
        // "Mark components in scope" checklist.
        //
        // Order is preserved from the existing selection rather than rebuilt from the board, so an
        // entry that has been curated by hand keeps its arrangement.
        // ###########################################################################################
        // ###########################################################################################
        // Narrows an entry's in-scope AND completed lists together after its area has been resized,
        // returning both. Completed is narrowed to the SCOPE THAT REMAINS rather than to the raw
        // area scope, which is what enforces the invariant the two lists depend on: a completed
        // component is always one the entry still covers.
        //
        // Narrowing them independently against the area looks equivalent and is not. A label the
        // user had unticked from the scope but which the area still touches would survive in
        // completed - claiming a component was done for an entry that does not include it. Doing it
        // in one place stops the two calls being written the wrong way round at a call site.
        // ###########################################################################################
        public static (List<string> InScope, List<string> Completed) NarrowEntryToScope(
            IReadOnlyList<string>? selectedBoardLabels,
            IReadOnlyList<string>? completedBoardLabels,
            IReadOnlyCollection<string> boardLabelsInScope)
        {
            var inScope = NarrowSelectionToScope(selectedBoardLabels, boardLabelsInScope);
            var completed = NarrowSelectionToScope(completedBoardLabels, inScope);

            return (inScope, completed);
        }

        public static List<string> NarrowSelectionToScope(
            IReadOnlyList<string>? selectedBoardLabels,
            IReadOnlyCollection<string> boardLabelsInScope)
        {
            if (selectedBoardLabels == null || selectedBoardLabels.Count == 0)
            {
                return new List<string>();
            }

            // Case-insensitive, matching every other board-label comparison in the app - a label
            // stored as "u7" must still be recognised as in scope when the rects key it as "U7".
            var scope = new HashSet<string>(
                boardLabelsInScope ?? Array.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var kept = new List<string>(selectedBoardLabels.Count);

            foreach (string? label in selectedBoardLabels)
            {
                // A null or blank entry identifies no component, so it can never be in scope and is
                // dropped rather than carried forward as an unmatchable label.
                string trimmed = label?.Trim() ?? string.Empty;

                if (trimmed.Length > 0 && scope.Contains(trimmed))
                {
                    kept.Add(label!);
                }
            }

            return kept;
        }

        // ###########################################################################################
        // Returns true when the provided board has at least one component explicitly tagged as PAL or NTSC.
        // This is what decides whether the region switch is offered at all for a board.
        // ###########################################################################################
        public static bool HasExplicitRegionComponents(BoardData? boardData)
        {
            if (boardData == null)
                return false;

            return boardData.Components.Any(component =>
                string.Equals(component.Region?.Trim(), "PAL", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(component.Region?.Trim(), "NTSC", StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Returns true when the supplied path points at a supported modern KiCad raw file.
        // ###########################################################################################
        public static bool IsSupportedKiCadRawFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            string extension = Path.GetExtension(path.Trim());

            return string.Equals(extension, ".kicad_pcb", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".kicad_pro", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".kicad_sch", StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Appends protected-file count information to a sync banner message when applicable.
        // ###########################################################################################
        public static string BuildSyncBannerText(string message, int protectedFilesCount)
        {
            return protectedFilesCount > 0
                ? $"{message}; protected contribution related files are [{protectedFilesCount}]"
                : message;
        }
    }
}
