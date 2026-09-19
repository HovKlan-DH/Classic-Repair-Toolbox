using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // The component checklist: "Mark components in scope" and "Mark components completed" rows,
    // their select-all/none actions and row click-to-toggle, and the completed-count summary. See
    // WorklogEntryEditorWindow.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class WorklogEntryEditorWindow
    {
        // "Mark components in scope". Populated only when the caller supplies the components -
        // this window has no board data and no highlight rectangles of its own, so it cannot work
        // out which components an area touches; see InitializeComponentScope.
        private readonly ObservableCollection<WorklogEntryComponentRow> thisComponentRows = new();

        // "Mark components completed" - one row per component currently TICKED in the scope list
        // above, carrying whether it has been done. A separate collection rather than a second flag
        // on the scope rows, because the two lists hold different sets: the scope list offers every
        // component the area touches, this one offers only those the user put in scope.
        private readonly ObservableCollection<WorklogEntryComponentRow> thisCompletedComponentRows = new();

        // Whether the caller supplied a scope at all. Distinct from "the list is empty": an area
        // that genuinely touches nothing shows "No components in this area", whereas an unknown
        // scope hides the section and, crucially, leaves the entry's saved ComponentLabels alone
        // rather than overwriting them with an empty list.
        private bool thisHasComponentScope;

        // ###########################################################################################
        // Supplies the "Mark components in scope" checklist. Called by the opener straight after
        // Initialize, because working out which components an entry's area touches needs the board
        // data and the per-schematic highlight rectangles - both of which live in TabSchematics,
        // not here. This window just renders what it is given and reports back the ticked rows.
        //
        // Each row starts ticked if its label is already in the entry's saved ComponentLabels, so
        // reopening an entry shows the choice the user made last time rather than re-ticking
        // everything.
        //
        // tickAll overrides that for a NEW entry, where there is no saved selection to restore and
        // every row would otherwise start unticked. The user drew the area around these components,
        // so all of them in scope is the right starting point and unticking one is quicker than
        // ticking eight - which is what the quick "New fault" card this flow replaced did too.
        // ###########################################################################################
        public void InitializeComponentScope(
            IReadOnlyList<(string BoardLabel, string DisplayName)> componentsInScope,
            bool tickAll = false)
        {
            // Populating the checklist is not an edit, and building the rows drives the CheckBox
            // bindings - without a guard the window opened with Save already enabled, making every
            // entry look modified before the user had touched anything.
            //
            // The guard is NOT lowered here. Initialize raises it and posts the lift at Background
            // priority so that its own TextBox assignments' queued TextChanged events run while it
            // is still up; lowering it synchronously at the end of this method - which the caller
            // runs immediately after Initialize - put it back down before those queued events
            // arrived, so they called MarkDirty and set thisIsDirty on an untouched window.
            //
            // (That was masked only because Initialize's posted job later reset thisIsDirty, so the
            // Save button looked right while the flag was transiently wrong. Verified: after the
            // normal-priority jobs ran, dirty was true and Save was enabled.)
            //
            // Leaving the lift to Initialize's single posted job also makes the flag's lifetime the
            // same whether or not a scope was supplied - it previously differed depending on
            // whether the caller could work one out.
            this.thisIsInitializing = true;

            this.PopulateComponentScope(componentsInScope, tickAll);

            this.EditorSaveButton.IsEnabled = false;
        }

        private void PopulateComponentScope(
            IReadOnlyList<(string BoardLabel, string DisplayName)> componentsInScope,
            bool tickAll = false)
        {
            this.thisHasComponentScope = true;
            this.thisComponentRows.Clear();

            var alreadySelected = new HashSet<string>(
                this.thisEntry.ComponentLabels ?? new List<string>(),
                StringComparer.OrdinalIgnoreCase);

            foreach (var component in componentsInScope)
            {
                this.thisComponentRows.Add(new WorklogEntryComponentRow
                {
                    BoardLabel = component.BoardLabel,
                    DisplayName = component.DisplayName,
                    IsChecked = tickAll || alreadySelected.Contains(component.BoardLabel)
                });
            }

            this.EditorComponentScopePanel.IsVisible = true;

            // Decided here, beside the scope panel's own visibility, rather than inside the count
            // helper - that ran on every checkbox tick, re-asserting the panel's visibility as a
            // side effect of updating a label and silently overriding anything that might later
            // want to hide it.
            this.EditorComponentCompletedPanel.IsVisible = true;
            this.EditorComponentCountText.Text = $"{this.thisComponentRows.Count(r => r.IsChecked)} of {this.thisComponentRows.Count} selected";
            this.EditorNoComponentsText.IsVisible = this.thisComponentRows.Count == 0;

            this.PopulateCompletedComponentRows();
        }

        // ###########################################################################################
        // Builds the completed checklist for a freshly opened entry: one row per in-scope component,
        // ticked if the entry has it saved as completed.
        //
        // Separate from RefreshCompletedComponentRows because the source of the ticks differs. That
        // one carries ticks forward from the rows already on screen, which is right for a live scope
        // edit; on open there are no such rows, and the saved list is the only truth.
        // ###########################################################################################
        private void PopulateCompletedComponentRows()
        {
            // On open there are no rows to carry ticks from, so the saved list is the only truth.
            this.RebuildCompletedComponentRows(this.thisEntry.CompletedComponentLabels ?? new List<string>());
        }

        // ###########################################################################################
        // Rebuilds the "Mark components completed" checklist from whatever is currently TICKED in
        // the scope list above. Called after every change to that list, so the two can never
        // disagree about which components the entry covers.
        //
        // Existing completed ticks are preserved across the rebuild, keyed by board label - the
        // rows are recreated but the user's progress is not thrown away by an unrelated scope edit.
        //
        // Two rules that fall out of this, both deliberate:
        //   - a component newly ticked INTO scope appears here UNTICKED. It is work still to do,
        //     which is the whole point of the list; arriving pre-ticked would claim it was already
        //     done and quietly overstate progress.
        //   - a component unticked OUT of scope loses its completed state entirely. It is no longer
        //     part of the entry, so a remembered "done" flag would be about a component the entry
        //     does not cover, and would resurface if the label was ever re-added.
        // ###########################################################################################
        private void RefreshCompletedComponentRows()
        {
            // Ticks carried across from the rows already on screen, so progress survives a rebuild
            // triggered by an unrelated scope edit. Read BEFORE the rebuild clears them.
            this.RebuildCompletedComponentRows(
                this.thisCompletedComponentRows.Where(r => r.IsChecked).Select(r => r.BoardLabel));
        }

        // ###########################################################################################
        // The single rebuild both entry points share: one row per component currently ticked in the
        // scope list, ticked if its label is in the given set.
        //
        // The two callers differ only in where that set comes from - the rows on screen for a live
        // scope edit, the saved list on open - so they were one copy-pasted body apart, which is how
        // a later change to the row shape or the comparer ends up applied to only one of them.
        // ###########################################################################################
        private void RebuildCompletedComponentRows(IEnumerable<string> tickedLabels)
        {
            var ticked = new HashSet<string>(tickedLabels, StringComparer.OrdinalIgnoreCase);

            this.thisCompletedComponentRows.Clear();

            foreach (var row in this.thisComponentRows.Where(r => r.IsChecked))
            {
                this.thisCompletedComponentRows.Add(new WorklogEntryComponentRow
                {
                    BoardLabel = row.BoardLabel,
                    DisplayName = row.DisplayName,
                    IsChecked = ticked.Contains(row.BoardLabel)
                });
            }

            this.UpdateCompletedComponentSummary();
        }

        // The count reads as progress ("3 of 8 completed") rather than as a bare total - the list
        // exists to answer "how much is left", and a total alone does not.
        private void UpdateCompletedComponentSummary()
        {
            int total = this.thisCompletedComponentRows.Count;
            int done = this.thisCompletedComponentRows.Count(r => r.IsChecked);

            this.EditorCompletedCountText.Text = $"{done} of {total} completed";
            this.EditorNoCompletedText.IsVisible = total == 0;
        }

        private void OnEditorSelectAllCompletedClick(object? sender, RoutedEventArgs e)
        {
            foreach (var row in this.thisCompletedComponentRows)
            {
                row.IsChecked = true;
            }

            this.UpdateCompletedComponentSummary();
            this.MarkDirty();
        }

        private void OnEditorSelectNoneCompletedClick(object? sender, RoutedEventArgs e)
        {
            foreach (var row in this.thisCompletedComponentRows)
            {
                row.IsChecked = false;
            }

            this.UpdateCompletedComponentSummary();
            this.MarkDirty();
        }

        private void OnEditorCompletedRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is Control control && control.DataContext is WorklogEntryComponentRow row)
            {
                row.IsChecked = !row.IsChecked;
                this.UpdateCompletedComponentSummary();
                this.MarkDirty();
            }
        }

        // ###########################################################################################
        // "All" / "None" bulk links, and whole-row click-to-toggle - the same interactions the quick
        // card's checklist offers. Each marks the window dirty so the Save button enables, since
        // changing the scope is a real edit to the entry.
        // ###########################################################################################
        private void OnEditorSelectAllComponentsClick(object? sender, RoutedEventArgs e)
        {
            foreach (var row in this.thisComponentRows)
            {
                row.IsChecked = true;
            }

            this.RefreshCompletedComponentRows();
            this.MarkDirty();
        }

        private void OnEditorSelectNoneComponentsClick(object? sender, RoutedEventArgs e)
        {
            foreach (var row in this.thisComponentRows)
            {
                row.IsChecked = false;
            }

            this.RefreshCompletedComponentRows();
            this.MarkDirty();
        }

        // The checkbox and both labels are IsHitTestVisible="False", so this Border handler is the
        // only thing that sees the click - which is what makes the whole row a hit target.
        private void OnEditorComponentRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is Control control && control.DataContext is WorklogEntryComponentRow row)
            {
                row.IsChecked = !row.IsChecked;
                this.RefreshCompletedComponentRows();
                this.MarkDirty();
            }
        }
    }
}
