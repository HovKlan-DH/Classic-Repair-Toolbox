using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // SEVERAL COMPONENTS PICKED IN THE LIST, AND THE PICK THE OTHER TABS HAND OVER (owner requests, 2026-11-08: "I want
    // to be able to mark multiple components", and "always keep the same selection for the left-side panel"). What the
    // list has selected, shown selected again as the pick moves, opened together in the pane or the table, and what the
    // other tabs' lists had selected as the tab is entered - "Mark all" there shown, never opened. In a part of its own
    // since 2026-11-09 (code review: TabDrafts.Workspace.cs had grown past 1,500 lines).
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // ###########################################################################################
        // The labels selected in the list, in its order - several while Ctrl or Shift picked more than one.
        // ###########################################################################################
        private List<string> SelectedInList()
        {
            if (this.thisComponentList?.SelectedItems is not { } selected)
            {
                return [];
            }

            // A set, so each row is one look-up (second code review, 2026-11-09: a walk of the selected per row, at every click).
            var selectedNow = new HashSet<object>(selected.Cast<object>(), ReferenceEqualityComparer.Instance);

            return this.thisComponentItems
                .Where(item => item.BoardLabel.Length > 0 && selectedNow.Contains(item))
                .Select(item => item.BoardLabel)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // The several picked shown selected in the list - changed only where it differs, so the list keeps its place.
        private void ShowSeveralSelected(ListBox list)
        {
            if (this.thisComponentKeyTimer?.IsEnabled == true || list.SelectedItems is not { } selected)
            {
                return;
            }

            // Sets, so each row is one look-up (second code review, 2026-11-09: "Mark all"'s 400 components were 160,000
            // comparisons twice over at every refresh of the left column).
            var picked = new HashSet<string>(this.thisPickedSeveral, StringComparer.OrdinalIgnoreCase);
            List<ComponentListItem> wanted = this.thisComponentItems.Where(item => picked.Contains(item.BoardLabel)).ToList();
            var selectedNow = new HashSet<object>(selected.Cast<object>(), ReferenceEqualityComparer.Instance);

            if (wanted.Count == selected.Count && wanted.All(selectedNow.Contains))
            {
                return;
            }

            this.thisFollowingComponentSelection = true;

            try
            {
                selected.Clear();

                foreach (ComponentListItem item in wanted)
                {
                    selected.Add(item);
                }
            }
            finally
            {
                this.thisFollowingComponentSelection = false;
            }
        }

        // ###########################################################################################
        // Several picked in the list: the pane on all of them together, or in "Table" the sheet of the entry open showing
        // their rows alone. Picked again as they are, nothing is read again.
        // ###########################################################################################
        private void OpenPickedComponents(List<string> labels)
        {
            if (this.IsUpdateRequiredShown)
            {
                this.ShowNavSelection();
                return;
            }

            this.thisPickedComponent = labels[0];
            this.thisPickedSeveral = labels;

            int generation = ++this.thisPickGeneration;

            if (!this.thisTableMode && this.IsPaneOnPicked())
            {
                return;
            }

            this.thisPicksOpening++;

            this.RunGuarded(async () =>
            {
                try
                {
                    if (generation != this.thisPickGeneration)
                    {
                        return;
                    }

                    if (this.thisTableMode)
                    {
                        ContributeEntry entry = this.EntryOpen() is { } shown && ContributeEntries.PageOf(shown) is not null
                            ? shown
                            : ContributeEntries.OfPage(this.ComponentPane.Page);

                        await this.ShowEntryInTableAsync(entry, takeFocus: false);
                    }
                    else
                    {
                        await this.OpenSeveralAsync(labels, () => generation == this.thisPickGeneration);
                    }
                }
                finally
                {
                    this.thisPicksOpening--;
                }

                this.ShowNavSelection();
            }, "Open components");
        }

        // ###########################################################################################
        // *** THE COMPONENTS SELECTED ON THE OTHER TABS ARE THE ONES PICKED HERE (owner request, 2026-11-08: "always keep the
        // same selection for the left-side panel "Category" and "Component" lists, meaning it should not change selections
        // when shifting between e.g. "Schematics" tab and "Contribute" tab"; Main.FollowContributeTabInLists). *** Handed
        // over as the tab is shown: the same as picked here changes nothing. Otherwise they are picked - those this list has -
        // and opened where a component's page or its table is what shows; else only marked in the list. None: none picked.
        // ###########################################################################################
        // `everyRowOfTheirs`: every row of the other tabs' list is selected - what "Mark all" there does.
        internal void TakePickFromOtherTabs(IReadOnlyList<string> labels, bool everyRowOfTheirs = false)
        {
            ArgumentNullException.ThrowIfNull(labels);

            var given = new HashSet<string>(labels, StringComparer.OrdinalIgnoreCase);

            List<string> wanted = this.thisComponentItems
                .Where(item => item.BoardLabel.Length > 0 && given.Contains(item.BoardLabel))
                .Select(item => item.BoardLabel)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (wanted.SequenceEqual(this.PickedLabels, StringComparer.OrdinalIgnoreCase))
            {
                return;
            }

            if (wanted.Count == 0)
            {
                this.LetGoOfPickedComponent();
                return;
            }

            bool componentShown = this.EntryOpen() is { } entry && ContributeEntries.PageOf(entry) is not null;

            // ###########################################################################################
            // *** "MARK ALL" ON THE OTHER TABS IS SHOWN HERE, NEVER OPENED (code review, 2026-11-09). *** Every component the
            // list shows, selected there to see every highlight, opened the pane on all of them - and the next field typed
            // went to every component of the board. "Mark all" is off on this tab for that reason: the list shows them
            // selected, as the lists are one, and a component's page says to click one, or pick several, to edit them -
            // whichever page is shown now or later (IsEveryTakenOver; second code review, 2026-11-09: a page picked
            // afterwards opened them all). The table is no such risk - it shows every row of the sheet, as with none picked.
            // Told by the other list itself, every row of it selected (third code review, 2026-11-09: compared with this list's
            // count, which lets through components with no category theirs does not, it missed - and opened them all).
            // ###########################################################################################
            if (wanted.Count > 1 && everyRowOfTheirs)
            {
                // The table, as with none picked, narrowed to nobody (TableLabels; second code review, 2026-11-09: narrowed to
                // the labels both lists have, it hid the components with no category theirs never lists - those whose error
                // stops Submit - and took the rows' grips and a paste's new rows away).
                this.thisEveryTakenOver = wanted;
                this.ShowEveryComponentSelected(wanted);
                return;
            }

            if (componentShown && wanted.Count > 1)
            {
                this.OpenPickedComponents(wanted);
                return;
            }

            if (componentShown && this.thisComponentItems.FirstOrDefault(item => string.Equals(item.BoardLabel, wanted[0], StringComparison.OrdinalIgnoreCase)) is { } one)
            {
                this.OpenPickedComponent(one);
                return;
            }

            this.thisPickedComponent = wanted[0];
            this.thisPickedSeveral = wanted.Count > 1 ? wanted : [];
            this.ShowNavSelection();
        }

        // The pick "Mark all" on the other tabs handed over - the very list, so any other pick, a new list, is not it.
        private List<string>? thisEveryTakenOver;

        // Whether the components picked are "Mark all"'s from the other tabs, which no page opens.
        private bool IsEveryTakenOver => this.thisPickedSeveral.Count > 1 && ReferenceEquals(this.thisPickedSeveral, this.thisEveryTakenOver);

        // The components the table is narrowed to - none for "Mark all"'s, which shows every row, as with none picked.
        private IReadOnlyList<string> TableLabels => this.IsEveryTakenOver ? [] : this.PickedLabels;

        // Every component picked, none open: the page the pane was on says so (ShowNoComponentPage), and a pick still opening
        // is overtaken.
        private void ShowEveryComponentSelected(List<string> labels)
        {
            this.thisPickedComponent = labels[0];
            this.thisPickedSeveral = labels;
            this.thisPickGeneration++;

            if (this.IsComponentOpen && this.ComponentPane.BoardPart is null && !this.ComponentPane.IsNewComponentForm)
            {
                this.thisNoComponentPage = this.ComponentPane.Page;
                this.CloseComponent();
            }

            this.ShowWorkspace();
        }

        // The components picked here, for the other tabs' list as this tab is left.
        internal IReadOnlyList<string> PickedForOtherTabs => this.PickedLabels;
    }
}
