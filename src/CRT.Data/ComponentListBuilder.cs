using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
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
    //
    // Problems (owner request, 2026-11-04) is what the draft's checks found on the component, said on its row
    // while the list is the Contribute tab's. Changed on the row as it stands: a row made again lets go of
    // the selection, and the list scrolls back to it - the panes' every save would do that.
    // ###########################################################################################
    public sealed class ComponentListItem : INotifyPropertyChanged
    {
        private BoardProblemCounts thisProblems;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string DisplayText { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string SelectionKey { get; init; } = string.Empty;
        public bool IsDrafted { get; init; }

        public BoardProblemCounts Problems
        {
            get => this.thisProblems;
            set
            {
                if (this.thisProblems == value)
                {
                    return;
                }

                this.thisProblems = value;

                foreach (string shown in (string[])[nameof(this.HasErrors), nameof(this.ErrorsText), nameof(this.ErrorsTip), nameof(this.HasWarnings), nameof(this.WarningsText), nameof(this.WarningsTip)])
                {
                    this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(shown));
                }
            }
        }

        public bool HasErrors => this.thisProblems.Errors > 0;

        public string ErrorsText => this.thisProblems.Errors.ToString(CultureInfo.InvariantCulture);

        public string ErrorsTip => this.thisProblems.Errors == 1 ? "1 error on this component" : $"{this.ErrorsText} errors on this component";

        public bool HasWarnings => this.thisProblems.Warnings > 0;

        public string WarningsText => this.thisProblems.Warnings.ToString(CultureInfo.InvariantCulture);

        public string WarningsTip => this.thisProblems.Warnings == 1 ? "1 warning on this component" : $"{this.WarningsText} warnings on this component";

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
        // *** WHICH ROWS OF A LIST MADE AGAIN ARE THE COMPONENTS HELD SELECTED - every restore of Main's selection goes by it
        // (Main.SelectInComponentList). *** A row whose key is held, and the rows of the LABEL of a key no row has any more:
        // a key holds the component's names and its region, so its names changed on the Contribute tab since it was held
        // (code review, 2026-11-09), or the list is of the other region (owner report, 2026-11-13: "when I ... have selected
        // one component in the component list, and I then change region, then it deselects that component") - a component
        // told apart per region has a row, and a key, in each. It is the same component either way. One with no row in the
        // list now has nothing to select. In ONE place since code review, 2026-11-13: the region's fallback stood beside
        // the region switch alone, and the selection held across a load (Main._selectionToKeep) lost the component.
        // ###########################################################################################
        public static List<int> RowsToSelect(IReadOnlyList<ComponentListItem> rows, IEnumerable<string> selectedKeys)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(selectedKeys);

            var held = new HashSet<string>(selectedKeys, StringComparer.OrdinalIgnoreCase);
            var present = new HashSet<string>(rows.Select(row => row.SelectionKey), StringComparer.OrdinalIgnoreCase);
            var labelsGone = new HashSet<string>(
                held.Where(key => !present.Contains(key)).Select(ComponentListBuilder.LabelOf).Where(label => label.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            var indexes = new List<int>();

            for (int i = 0; i < rows.Count; i++)
            {
                if (held.Contains(rows[i].SelectionKey) || (labelsGone.Count > 0 && labelsGone.Contains(rows[i].BoardLabel)))
                {
                    indexes.Add(i);
                }
            }

            return indexes;
        }

        // ###########################################################################################
        // A ROW'S KEY (ComponentListItem.SelectionKey): its label, friendly name, technical name and region, joined by U+001F
        // so a field holding a pipe or a space keeps it unique. Made by BuildComponentItems and read apart HERE alone (code
        // review, 2026-11-13: three places split it, each with its own idea of the layout).
        // ###########################################################################################
        private const char KeySeparator = '\u001F';

        public static string LabelOf(string? selectionKey) => (selectionKey ?? string.Empty).Split(ComponentListBuilder.KeySeparator)[0];

        // Its friendly and technical names - empty for a key without them.
        public static (string Friendly, string Technical) NamesOf(string? selectionKey)
        {
            string[] parts = (selectionKey ?? string.Empty).Split(ComponentListBuilder.KeySeparator);

            return (parts.Length > 1 ? parts[1] : string.Empty, parts.Length > 2 ? parts[2] : string.Empty);
        }

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

                if (!ComponentListBuilder.MatchesSearch(displayString, searchTerm))
                    continue;

                string boardLabel = component.BoardLabel?.Trim() ?? string.Empty;

                items.Add(new ComponentListItem
                {
                    BoardLabel = boardLabel,
                    DisplayText = displayString,
                    SelectionKey = string.Join(ComponentListBuilder.KeySeparator,
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
        // The search's one rule: every space-separated term of `searchTerm` in an item's display text,
        // capitals ignored - a blank search matches everything. BuildComponentItems filters by it, and
        // a list already built is narrowed by it the same way (code review, 2026-10-30: the Contribute
        // tab built its list twice per refresh, once only to count).
        // ###########################################################################################
        public static bool MatchesSearch(string displayText, string? searchTerm)
        {
            if (string.IsNullOrWhiteSpace(searchTerm))
                return true;

            foreach (var term in searchTerm.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                if (displayText.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                    return false;
            }

            return true;
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
    }
}
