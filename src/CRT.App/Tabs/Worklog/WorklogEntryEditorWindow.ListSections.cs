using Avalonia.Controls;
using Avalonia.Input;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // Shared collapsible list-section infrastructure the Links/Comments/WorkDoneItems sections
    // (and the component checklist) all use: expand/collapse state, persistence, empty-state text
    // and item-count labels. See WorklogEntryEditorWindow.axaml.cs for the file map of the whole
    // partial class.
    // ###########################################################################################
    public partial class WorklogEntryEditorWindow
    {
        // ###########################################################################################
        // Collapsible list sections.
        //
        // Each of the seven lists (Links, Work done, Comments, Components in scope, Components
        // completed, Photos, Files) has a header the user can click to fold its content away, so a
        // worklog with a long checklist and forty photos can be skimmed rather than scrolled past.
        //
        // Driven by one table rather than seven near-identical handlers: the sections differ only in
        // which controls they own, and duplicating the toggle logic per section is how one of them
        // eventually ends up with a subtly different rule.
        //
        // Collapsed state IS persisted, per entry, in entries.json - see PersistCollapsedSections.
        // It is a reading convenience rather than an edit, so it saves itself immediately instead of
        // waiting for "Update worklog", and it never marks the window dirty.
        //
        // (An earlier draft of this comment said the opposite. Only the collapsed sections are
        // stored, so an absent key means "expanded" and an entry written before the field existed -
        // or one never folded - opens with everything showing.)
        // ###########################################################################################
        private sealed class WorklogListSection
        {
            public required TextBlock Icon { get; init; }

            // The controls folded away, and shown again unconditionally when the section expands.
            public required IReadOnlyList<Control> Body { get; init; }

            // The section's "No links added" line, if it has one. Kept apart from Body because
            // whether it belongs on screen depends on whether the list is EMPTY, not on whether the
            // section is open - showing it with the rest would put "No links added" above a list of
            // links. Expanding therefore asks the refresh methods to restore it.
            public Control? EmptyState { get; init; }

            public bool IsExpanded { get; set; } = true;
        }

        private readonly Dictionary<string, WorklogListSection> thisListSections = new(StringComparer.Ordinal);

        // fa-regular square-plus / square-minus. Read out of the shipped OTF rather than from
        // memory: the Free Regular face is a 362-glyph subset, so a codepoint that exists in Solid
        // is often absent here and renders as a blank box with nothing failing.
        private const string ExpandIconGlyph = "";

        private const string CollapseIconGlyph = "";

        private void InitializeListSections()
        {
            this.thisListSections["EditorLinksHeader"] = new WorklogListSection
            {
                Icon = this.EditorLinksHeaderIcon,
                Body = new Control[] { this.EditorLinksList },
                EmptyState = this.EditorNoLinksText,
            };

            this.thisListSections["EditorWorkDoneHeader"] = new WorklogListSection
            {
                Icon = this.EditorWorkDoneHeaderIcon,
                Body = new Control[] { this.EditorWorkDoneList },
                EmptyState = this.EditorNoWorkDoneText,
            };

            this.thisListSections["EditorCommentsHeader"] = new WorklogListSection
            {
                Icon = this.EditorCommentsHeaderIcon,
                Body = new Control[] { this.EditorCommentsList },
                EmptyState = this.EditorNoCommentsText,
            };

            // The checklists fold their whole bordered box, not the ItemsControl inside it - the
            // border is the visible extent of the list, so leaving it behind would collapse the
            // rows into an empty frame rather than out of the way.
            this.thisListSections["EditorComponentScopeHeader"] = new WorklogListSection
            {
                Icon = this.EditorComponentScopeHeaderIcon,
                Body = new Control[] { this.EditorComponentScopeBody },
            };

            this.thisListSections["EditorComponentCompletedHeader"] = new WorklogListSection
            {
                Icon = this.EditorComponentCompletedHeaderIcon,
                Body = new Control[] { this.EditorComponentCompletedBody },
            };

            this.thisListSections["EditorPhotosHeader"] = new WorklogListSection
            {
                Icon = this.EditorPhotosHeaderIcon,
                Body = new Control[] { this.EditorPhotosList },
                EmptyState = this.EditorNoPhotosText,
            };

            this.thisListSections["EditorFilesHeader"] = new WorklogListSection
            {
                Icon = this.EditorFilesHeaderIcon,
                Body = new Control[] { this.EditorFilesList },
                EmptyState = this.EditorNoFilesText,
            };

            foreach (var section in this.thisListSections.Values)
            {
                ApplyListSectionState(section);
            }
        }

        private void OnListHeaderTogglePointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Control { Tag: string key })
                return;

            this.SetListSectionExpanded(key, !this.IsListSectionExpanded(key));
        }

        // ###########################################################################################
        // Opens or folds one section and writes the change straight to disk.
        //
        // Persisted immediately rather than with Save, matching the sub-lists: which sections are
        // folded is a reading preference, not an edit to the worklog, so it must not enable the
        // "Update worklog" button or be discardable with Cancel. Nor may it be lost by closing the
        // window the way it was opened.
        // ###########################################################################################
        private void SetListSectionExpanded(string key, bool isExpanded)
        {
            if (!this.thisListSections.TryGetValue(key, out var section))
                return;

            if (section.IsExpanded == isExpanded)
                return;

            section.IsExpanded = isExpanded;
            ApplyListSectionState(section);

            if (section.IsExpanded)
            {
                this.RefreshListSectionEmptyStates();
            }

            this.PersistCollapsedSections();
        }

        // ###########################################################################################
        // Expands a section that is folded, used when something is ADDED to it - a new comment that
        // lands in a collapsed list would otherwise appear to have gone nowhere.
        //
        // Only ever opens, never closes: a user who has a section open and adds to it must not have
        // it fold underneath them.
        // ###########################################################################################
        private void EnsureListSectionExpanded(string key) => this.SetListSectionExpanded(key, true);

        // ###########################################################################################
        // Writes ONLY the fold state, by re-reading the stored record and putting the folds on that.
        //
        // Deliberately not PersistEntrySilently, which syncs every direct field first. Folding a
        // section is a reading convenience, not an edit - the Description, category, state and
        // "Show marked area" are the fields the user can still abandon with Cancel, and routing a
        // fold through the sub-list save path committed all of them to disk the moment a header was
        // clicked. Pressing Cancel afterwards then reported success with the abandoned edits live.
        //
        // Reading the stored record back rather than writing the working copy is what keeps those
        // pending edits out: the folds land on what is genuinely on disk. If the entry cannot be
        // read back - deleted from under the window - the fold is simply not persisted, which is
        // the right outcome for a preference.
        // ###########################################################################################
        private void PersistCollapsedSections()
        {
            if (this.thisIsInitializing)
                return;

            var collapsed = this.thisListSections
                .Where(pair => !pair.Value.IsExpanded)
                .Select(pair => pair.Key)
                .OrderBy(key => key, StringComparer.Ordinal)
                .ToList();

            // Kept on the working copy too, so a later Save does not write back the old folds.
            this.thisEntry.CollapsedSections = collapsed;

            var stored = WorklogManager.GetEntries(this.thisWorkbookId)
                .FirstOrDefault(entry => entry.Id == this.thisEntry.Id);

            if (stored == null)
                return;

            stored.CollapsedSections = collapsed;
            WorklogManager.UpdateEntry(this.thisWorkbookId, stored);
        }

        // ###########################################################################################
        // Applies the folds saved on the entry. Absent keys mean expanded, so an entry written before
        // this field existed - or one the user has never folded - opens with everything showing.
        // ###########################################################################################
        private void RestoreCollapsedSections()
        {
            var collapsed = new HashSet<string>(
                this.thisEntry.CollapsedSections ?? new List<string>(),
                StringComparer.Ordinal);

            foreach (var (key, section) in this.thisListSections)
            {
                section.IsExpanded = !collapsed.Contains(key);
                ApplyListSectionState(section);
            }
        }

        // ###########################################################################################
        // Re-applies every list's "No ... added" line from its row count and its section's fold
        // state. Called after expanding a section, because whether that line belongs on screen is a
        // property of the DATA - showing it along with the rest of the body would put "No links
        // added" above a list that has links in it.
        // ###########################################################################################
        private void RefreshListSectionEmptyStates()
        {
            this.EditorNoLinksText.IsVisible = this.thisLinkRows.Count == 0 && this.IsListSectionExpanded("EditorLinksHeader");
            this.EditorNoWorkDoneText.IsVisible = this.thisWorkDoneRows.Count == 0 && this.IsListSectionExpanded("EditorWorkDoneHeader");
            this.EditorNoCommentsText.IsVisible = this.thisCommentRows.Count == 0 && this.IsListSectionExpanded("EditorCommentsHeader");
            this.EditorNoPhotosText.IsVisible = this.thisPhotoRows.Count == 0 && this.IsListSectionExpanded("EditorPhotosHeader");
            this.EditorNoFilesText.IsVisible = this.thisFileRows.Count == 0 && this.IsListSectionExpanded("EditorFilesHeader");
        }

        // ###########################################################################################
        // Shows or hides a section's body and swaps its icon.
        //
        // The empty-state line is hidden on collapse but NOT shown on expand - whether it belongs
        // on screen depends on whether the list is empty. RefreshListSectionEmptyStates restores it
        // from the row counts after an expand.
        // ###########################################################################################
        private static void ApplyListSectionState(WorklogListSection section)
        {
            section.Icon.Text = section.IsExpanded ? CollapseIconGlyph : ExpandIconGlyph;

            foreach (var control in section.Body)
            {
                control.IsVisible = section.IsExpanded;
            }

            // Collapsing always hides the empty-state line. Expanding does NOT simply show it -
            // that is decided by whether the list has rows, so the caller refreshes it instead.
            if (section.EmptyState != null && !section.IsExpanded)
            {
                section.EmptyState.IsVisible = false;
            }
        }

        // ###########################################################################################
        // The item count shown beside a list's title. "none" rather than "0 items" for an empty
        // list: the section already carries a "No links added" line inside it, and a zero repeated
        // twice reads as noise.
        // ###########################################################################################
        private static string FormatItemCount(int count, string singular, string plural) =>
            count switch
            {
                0 => "none",
                1 => $"1 {singular}",
                _ => $"{count} {plural}",
            };

        private bool IsListSectionExpanded(string key) =>
            !this.thisListSections.TryGetValue(key, out var section) || section.IsExpanded;
    }
}
