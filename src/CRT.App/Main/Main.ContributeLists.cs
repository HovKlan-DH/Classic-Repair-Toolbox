using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE CONTRIBUTE TAB PICKS ITS COMPONENTS ON THE FAR LEFT, WHERE CRT'S LISTS STAND (owner request, 2026-11-03:
    // "why not utilize the existing lists already available in the left-side ... when I am in 'Contribute' tab,
    // then unselect everything in those lists (and remember it) and then have those lists work as selector for the
    // component. Only one component should be possible to select at a time").
    //
    // *** LISTS AND A BOX OF ITS OWN IN THAT PLACE, NEVER THE OTHER TABS' (code review, 2026-11-04). *** Showing the
    // tab changes only which are visible (FollowContributeTabInLists). It lent the one pair until then: each switch
    // swapped their rows, selection mode and words, which took a guard wherever a list was touched, records of
    // what they had held and scroll offsets saved by hand - and every review found another way for what the other
    // tabs had selected to get lost in it. The tab's show TabDrafts.CategoryItems and ComponentItems, which follow
    // the draft as it is edited - the board on screen is read again only once the pane closes. On them:
    //
    //   - EVERY CATEGORY IS TICKED THE FIRST TIME, AND WHAT IS DONE TO THE LISTS HERE STAYS (owner request,
    //     2026-11-03: "should select all in the 'Category' list ... but only the first time ... if I switch
    //     between contribute and schematics tab, then I do not want it to reset every time ... same for the
    //     component selected"): the categories UNTICKED and the words typed are kept per board for the run
    //     (ContributeListState) - unticked rather than ticked, so a category the draft has just made is ticked.
    //     Never saved as the board's filter, which is the other list's.
    //   - A COMPONENT PICKED IS THE TAB'S, one at a time ("Mark all" is off), and stays picked whatever else the
    //     tab shows (TabDrafts.thisPickedComponent).
    //   - THE FILTER BOX ONLY NARROWS THE LIST, selecting nothing.
    //   - "WITH ISSUES" COMES WITH THE LIST, there on no other tab (TabDrafts.Problems.cs).
    //   - "Clear" lets go of all of it: every category ticked, nothing typed, nothing narrowed, no component picked.
    //   - THE REGION BUTTONS are the tab's region too (PickRegion) - and follow "Board: Regions" as it is
    //     ticked, so they read what the tab reads.
    //
    // The other tabs' lists stay behind them as they were, their selection and scroll with them; back on either,
    // everything stands where it stood (owner request, 2026-11-04: "navigate to the exact same position"). The one
    // thing that reaches them is a board read again - what the pane saved - which drops a selection as any load
    // does, so what was selected as the tab was entered is held until such a load has shown it (_selectionToKeep).
    // Its rows stay on screen until the new ones come, and those come selected in one go (LetGoOfListsForLoad,
    // SelectInComponentList; owner report, 2026-11-05).
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // Whether the Contribute tab is shown - so its lists rather than the other tabs'.
        private bool _contributeListsShown;

        // The components the other tabs' list selects again when the board is read again, for the board they are of.
        private (string BoardKey, HashSet<string> ComponentKeys)? _selectionToKeep;

        // The board whose rows the other tabs' lists show - null with none.
        private string? _boardInOtherTabsLists;

        // The board what is ticked and typed in the tab's lists is filed under.
        private string _contributeListsBoard = string.Empty;

        // The board the tab's lists were last handed over for. Until a load ends it is not the one above: another
        // board chosen is filed under at once, while what they list and have ticked is still the board's before.
        private string _contributeListsReadFor = string.Empty;

        // The words the tab's list is narrowed by, and the wait after the last keystroke - as the other box has
        // (Main.ComponentPopup.cs), and TextChanged is posted here too.
        private string _contributeSearchInEffect = string.Empty;

        private DispatcherTimer? _contributeSearchTimer;

        // True while the ticks are set from code - no choice of anybody's.
        private bool _followingContributeTicks;

        internal bool ContributeListsShownForTests => this._contributeListsShown;

        // What the tab's lists say, for one board: kept for the run, so going back and forth between the tabs
        // resets nothing, and a board never looked at yet has every category ticked and nothing typed.
        private sealed class ContributeListState
        {
            public HashSet<string> Unticked { get; } = new(StringComparer.OrdinalIgnoreCase);

            public string Search { get; set; } = string.Empty;
        }

        private readonly Dictionary<string, ContributeListState> _contributeListStates = new(StringComparer.OrdinalIgnoreCase);

        private ContributeListState ContributeListStateOf(string boardKey)
        {
            if (!this._contributeListStates.TryGetValue(boardKey, out ContributeListState? state))
            {
                state = new ContributeListState();
                this._contributeListStates[boardKey] = state;
            }

            return state;
        }

        private void WireContributeLists()
        {
            this.ContributeCategoryListBox.ItemsSource = this.TabDrafts.CategoryItems;
            this.ContributeComponentListBox.ItemsSource = this.TabDrafts.ComponentItems;
            this.ContributeCategoryListBox.SelectionChanged += this.OnContributeCategorySelectionChanged;

            this.TabDrafts.UseComponentList(this.ContributeComponentListBox);
            this.TabDrafts.ComponentListsShown += this.OnContributeListsShown;
        }

        // ###########################################################################################
        // A new board renamed under "Board: Name" (code review, 2026-11-04): what the lists hold under its old
        // name goes with it, as everything else this machine files under the name does (BoardKeyMove).
        // ###########################################################################################
        internal void CarryListsToRenamedBoard(string from, string to)
        {
            if (this._contributeListStates.Remove(from, out ContributeListState? state))
            {
                this._contributeListStates[to] = state;
            }

            if (string.Equals(this._contributeListsBoard, from, StringComparison.OrdinalIgnoreCase))
            {
                this._contributeListsBoard = to;
            }

            if (string.Equals(this._contributeListsReadFor, from, StringComparison.OrdinalIgnoreCase))
            {
                this._contributeListsReadFor = to;
            }

            if (this._selectionToKeep is { } kept && string.Equals(kept.BoardKey, from, StringComparison.OrdinalIgnoreCase))
            {
                this._selectionToKeep = (to, kept.ComponentKeys);
            }

            if (string.Equals(this._boardInOtherTabsLists, from, StringComparison.OrdinalIgnoreCase))
            {
                this._boardInOtherTabsLists = to;
            }
        }

        // A board as a load leaves the lists, with no load: its data in and every list as it shows it.
        internal void UseBoardForTests(HardwareBoardEntry entry, BoardData board)
        {
            string boardKey = this.GetCurrentBoardKey();

            this._currentBoardData = board;
            this._suppressCategoryFilterSave = true;
            this.ShowListsOfBoard(board, boardKey, this._localRegion);
            this._suppressCategoryFilterSave = false;
            this.LetGoOfSelectionShown(boardKey);
            this.UseContributeListsOfBoardShown();
            this.TabDrafts.ShowBoard(entry, board, this._localRegion, null);
        }

        // A load begun, as far as the lists go: the board's data let go of, the lists as LetGoOfListsForLoad leaves them.
        internal void LoadBeginsForTests()
        {
            this.LetGoOfListsForLoad();
            this._currentBoardData = null;
        }

        // ###########################################################################################
        // The tab shown decides, from the tab control's selection change. A word waiting to be taken is taken
        // now, or it would be filed by a timer running behind the other tabs. The keyboard stays in "the filter
        // box": to the person using it this is one box.
        // ###########################################################################################
        private void FollowContributeTabInLists()
        {
            bool shown = this.ContributeTabItem is not null && ReferenceEquals(this.MainTabControl?.SelectedItem, this.ContributeTabItem);

            if (shown == this._contributeListsShown)
            {
                return;
            }

            if (shown)
            {
                this.HoldSelectionOfTheOtherTabs();
            }
            else if (!string.Equals(this.ContributeSearchTextBox.Text ?? string.Empty, this._contributeSearchInEffect, StringComparison.Ordinal))
            {
                this.ApplyContributeSearchNow();
            }

            TextBox boxShown = shown ? this.ContributeSearchTextBox : this.ComponentSearchTextBox;
            bool keyboardInTheBox = (shown ? this.ComponentSearchTextBox : this.ContributeSearchTextBox).IsFocused;

            this._contributeListsShown = shown;

            this.CategoryFilterListBox.IsVisible = !shown;
            this.ComponentFilterListBox.IsVisible = !shown;
            this.ComponentSearchTextBox.IsVisible = !shown;

            this.ContributeCategoryListBox.IsVisible = shown;
            this.ContributeComponentListBox.IsVisible = shown;
            this.ContributeSearchTextBox.IsVisible = shown;

            this.TabDrafts.UseIssuesBox(shown ? this.ComponentsWithIssuesCheckBox : null);

            if (!shown)
            {
                this.ComponentsWithIssuesCheckBox.IsVisible = false;
            }

            this.MarkAllComponentsButton.IsEnabled = !shown;
            this.UpdateRegionButtonsState();

            if (keyboardInTheBox)
            {
                boxShown.Focus();
            }
        }

        // ###########################################################################################
        // What the other tabs have selected as this one is entered, held for any board read again until one shows
        // it (ShowListsOfBoard). What is held already stays (code review, 2026-11-04: back on the tab before a load
        // ended, the selection was gone for good), unless somebody has chosen since, which lets go of it.
        //
        // A load under way has kept the rows of the board it reads again, so what the list has selected is still
        // there to hold (code review, 2026-11-05) - another board's took them, which leaves nothing of it, nor
        // anything for a word to be held over. Nothing selected is worth holding only under a word, which would
        // select all it finds.
        // ###########################################################################################
        private void HoldSelectionOfTheOtherTabs()
        {
            if (this.IsSelectionHeld)
            {
                return;
            }

            var keys = new HashSet<string>(
                this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>().Select(item => item.SelectionKey) ?? [],
                StringComparer.OrdinalIgnoreCase);

            if (this._currentBoardData is null && keys.Count == 0)
            {
                return;
            }

            this._selectionToKeep = keys.Count > 0 || !string.IsNullOrWhiteSpace(this.ComponentSearchTextBox.Text)
                ? (this.GetCurrentBoardKey(), keys)
                : null;
        }

        private bool IsSelectionHeld =>
            this._selectionToKeep is { } held && string.Equals(held.BoardKey, this.GetCurrentBoardKey(), StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // The lists have shown what was held. With the Contribute tab still shown it stays held, as the next
        // change made there reads the board again; once left the next load drops a selection as any load does.
        // Held for another board it is of no more use - that board's lists have gone from the screen.
        // ###########################################################################################
        private void LetGoOfSelectionShown(string boardKey)
        {
            if (!this._contributeListsShown ||
                this._selectionToKeep is { } held && !string.Equals(held.BoardKey, boardKey, StringComparison.OrdinalIgnoreCase))
            {
                this._selectionToKeep = null;
            }
        }

        private List<string> SelectedComponentLabels() =>
            this.ComponentFilterListBox.SelectedItems?
                .Cast<ComponentListItem>()
                .Select(item => item.BoardLabel)
                .Where(label => !string.IsNullOrEmpty(label))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? [];

        // ###########################################################################################
        // A load's last step: the board's rectangles are in, so what the other tabs' list has selected by now
        // lights up, whichever tab is shown (code review, 2026-11-04: decided when the list was shown, leaving the
        // Contribute tab during a load left a selection in the list with nothing lit on the picture).
        //
        // With nothing selected too (code review, 2026-11-05): the rows stay as a board is read again, so nothing
        // said the selection had gone, and the labels highlighted and the blink timer stood as they were.
        // ###########################################################################################
        internal void ShowHighlightsOfSelectionAfterLoad() =>
            this.TabSchematicsControl.UpdateHighlightsForComponents(this.SelectedComponentLabels());

        // ###########################################################################################
        // What the categories ticked and the box say now, kept for the board and handed to the tab - which
        // narrows its list by them. From a pick or a word typed, never from the lists being read again.
        //
        // *** NOT WHILE THE BOARD CHOSEN IS STILL BEING READ (code review, 2026-11-05). *** The lists then show the
        // board before, and the ticks read here were filed as the chosen board's - a word typed meanwhile was enough.
        // The word is kept already (TakeWaitingContributeWord), and the board shown reads both when it comes.
        // ###########################################################################################
        private void FollowContributeFilter()
        {
            if (!this.ContributeListsShowTheirBoard)
            {
                return;
            }

            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);
            var ticked = new HashSet<string>(this.ContributeCategoryListBox.SelectedItems?.Cast<string>() ?? [], StringComparer.OrdinalIgnoreCase);

            state.Unticked.Clear();
            state.Unticked.UnionWith(this.TabDrafts.CategoryItems.Where(category => !ticked.Contains(category)));

            this.TabDrafts.UseComponentFilter(state.Unticked, state.Search);
        }

        // A category ticked by hand - a category going as the list is read again is nobody's choice.
        private void OnContributeCategorySelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!this._followingContributeTicks && !this.TabDrafts.IsReadingComponentLists)
            {
                this.FollowContributeFilter();
            }
        }

        // ---- the tab's filter box ------------------------------------------------------------------

        private void ShowContributeSearchText(string text)
        {
            this._contributeSearchInEffect = text;
            this.ContributeSearchTextBox.Text = text;
        }

        // The filter runs 200 ms after the LAST keystroke, as the other box's does.
        private void OnContributeSearchTextChanged(object? sender, TextChangedEventArgs e)
        {
            if (string.Equals(this.ContributeSearchTextBox.Text ?? string.Empty, this._contributeSearchInEffect, StringComparison.Ordinal))
            {
                this._contributeSearchTimer?.Stop();
                return;
            }

            if (this._contributeSearchTimer == null)
            {
                this._contributeSearchTimer = new DispatcherTimer { Interval = Main.ComponentSearchDebounceInterval };
                this._contributeSearchTimer.Tick += (_, _) => this.ApplyContributeSearchNow();
            }

            this._contributeSearchTimer.Stop();
            this._contributeSearchTimer.Start();
        }

        private void ApplyContributeSearchNow()
        {
            this.TakeWaitingContributeWord();
            this.FollowContributeFilter();
        }

        // The word in the box, kept for the board it was typed for.
        private void TakeWaitingContributeWord()
        {
            this._contributeSearchTimer?.Stop();

            string word = this.ContributeSearchTextBox.Text ?? string.Empty;

            this._contributeSearchInEffect = word;
            this.ContributeListStateOf(this._contributeListsBoard).Search = word;
        }

        // ###########################################################################################
        // Another board chosen: a word still waiting stays with the one it was typed for, and the box says what
        // was left on this one. As soon as it is chosen, not once it is read: that can be seconds, and what is
        // typed meanwhile is typed for this board.
        // ###########################################################################################
        internal void FollowBoardInContributeLists()
        {
            string boardKey = this.GetCurrentBoardKey();

            if (string.Equals(boardKey, this._contributeListsBoard, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            this.TakeWaitingContributeWord();
            this._contributeListsBoard = boardKey;
            this.ShowContributeSearchText(this.ContributeListStateOf(boardKey).Search);
        }

        // ###########################################################################################
        // What narrows the tab's list, for the board about to be shown there. Only handed over: showing it reads
        // the list right after (code review, 2026-11-04: read here too, the whole left column was worked out
        // twice per load, the first time on the page and the filter of different boards). A word typed and not
        // taken yet is the contributor's, and is taken now, as its timer would.
        // ###########################################################################################
        private void UseContributeListsOfBoardShown()
        {
            this.FollowBoardInContributeLists();
            this.TakeWaitingContributeWord();

            this._contributeListsReadFor = this._contributeListsBoard;

            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);

            this.TabDrafts.KeepComponentFilter(state.Unticked, state.Search);
        }

        private bool ContributeListsShowTheirBoard =>
            string.Equals(this._contributeListsReadFor, this._contributeListsBoard, StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Every category listed ticked but those unticked by hand: the list changes in place as the draft
        // does, and a category that has just come is one nobody has unticked. Nothing here is a choice. The
        // lists read again while the board chosen is still coming keep their ticks: what is unticked on that
        // board says nothing of these categories.
        // ###########################################################################################
        private void ShowContributeTicks()
        {
            if (!this.ContributeListsShowTheirBoard)
            {
                return;
            }

            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);

            this._followingContributeTicks = true;

            try
            {
                for (int index = 0; index < this.TabDrafts.CategoryItems.Count; index++)
                {
                    bool ticked = !state.Unticked.Contains(this.TabDrafts.CategoryItems[index]);

                    if (ticked && !this.ContributeCategoryListBox.Selection.IsSelected(index))
                    {
                        this.ContributeCategoryListBox.Selection.Select(index);
                    }
                    else if (!ticked && this.ContributeCategoryListBox.Selection.IsSelected(index))
                    {
                        this.ContributeCategoryListBox.Selection.Deselect(index);
                    }
                }
            }
            finally
            {
                this._followingContributeTicks = false;
            }
        }

        // The lists were read again: the region buttons read what the tab reads ("Board: Regions" ticked
        // shows them at once), and the ticks follow the categories there now are.
        private void OnContributeListsShown(object? sender, EventArgs e)
        {
            this.UpdateRegionButtonsState();
            this.ShowContributeTicks();
        }

        // "Clear" on the tab: back to how it starts - nothing narrowed, and no component picked.
        private void ClearContributeLists()
        {
            this._contributeSearchTimer?.Stop();

            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);

            state.Unticked.Clear();
            state.Search = string.Empty;

            this.ShowContributeSearchText(string.Empty);

            this.TabDrafts.LetGoOfOnlyWithIssues();
            this.TabDrafts.UseComponentFilter(state.Unticked, state.Search);
            this.TabDrafts.LetGoOfPickedComponent();
        }

        // A press on "With issues" and its box, as on the "Labels visible" panel's rows: the box takes no pointer.
        private void OnComponentsWithIssuesRowClicked(object? sender, PointerPressedEventArgs e)
        {
            if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                this.ComponentsWithIssuesCheckBox.IsChecked = this.ComponentsWithIssuesCheckBox.IsChecked != true;
                e.Handled = true;
            }
        }

        // The "Open components" of "Overview": the keyboard to the filter over the list.
        internal void FocusComponentFilter() => this.ContributeSearchTextBox.Focus();

        // ---- the other tabs' lists, as a board is read ------------------------------------------------

        // ###########################################################################################
        // A load begins. *** THE BOARD ON SCREEN READ AGAIN KEEPS ITS ROWS UNTIL THE NEW ONES COME (owner report,
        // 2026-11-05: "it slowly populated the list of components"): *** it is read again at every picture tool
        // and every time the Contribute tab is left with a change made, and an empty list that fills again is
        // seen each time, the scroll position lost with it. Another board's rows go at once. Letting go reads as
        // a selection changed, no choice of anybody's, so what is held for this load to select again stays held.
        // ###########################################################################################
        //
        // *** THE ROWS KEPT TAKE NO CLICK UNTIL THE NEW ONES COME (code review, 2026-11-05). *** A category ticked
        // or a row selected meanwhile looked as if it had worked, and was gone as the load ended without a word.
        // They look the same - greyed out, the flicker the kept rows remove would be back - and take clicks again
        // with the new rows, or as the load ends however it ends (LoadSelectedBoardAsync).
        // ###########################################################################################
        internal void LetGoOfListsForLoad()
        {
            if (this._boardInOtherTabsLists is null ||
                !string.Equals(this._boardInOtherTabsLists, this.GetCurrentBoardKey(), StringComparison.OrdinalIgnoreCase))
            {
                this.LetGoOfOtherTabsLists();
                return;
            }

            this.OtherTabsListsTakeClicks(false);
        }

        // The board that could not be read after all, or another board chosen: its rows go.
        private void LetGoOfOtherTabsLists()
        {
            (string BoardKey, HashSet<string> ComponentKeys)? held = this._selectionToKeep;

            this._boardInOtherTabsLists = null;
            this.CategoryFilterListBox.ItemsSource = null;
            this.ComponentFilterListBox.ItemsSource = null;
            this.OtherTabsListsTakeClicks(true);

            this._selectionToKeep = held;
        }

        private void OtherTabsListsTakeClicks(bool take)
        {
            this.CategoryFilterListBox.IsHitTestVisible = take;
            this.ComponentFilterListBox.IsHitTestVisible = take;
        }

        // ###########################################################################################
        // *** ROWS SELECTED FROM CODE ARE SELECTED IN ONE GO, AND THE LIST STAYS PUT (owner report, 2026-11-05: "it
        // slowly populated the list of components in the left-side, and then marked all components"; measured:
        // 276 scrolls for 300 rows). *** Row by row, each moved the list's anchor and asked to be scrolled to it,
        // a layout pass apiece, so the list was seen running down to its last selected row at every load. Its
        // AutoScrollToSelectedItem is off in the markup for the same reason - Avalonia also scrolls as a list
        // comes back into view - so the one selection that must be seen asks for it (SelectComponentByBoardLabel).
        // ###########################################################################################
        private void SelectInComponentList(IReadOnlyList<ComponentListItem> rows, HashSet<string> selectionKeys)
        {
            this.ComponentFilterListBox.Selection.BeginBatchUpdate();

            try
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    if (selectionKeys.Contains(rows[i].SelectionKey))
                    {
                        this.ComponentFilterListBox.Selection.Select(i);
                    }
                }
            }
            finally
            {
                this.ComponentFilterListBox.Selection.EndBatchUpdate();
            }
        }

        // ###########################################################################################
        // The lists of a board, as it is shown: its categories as the board's filter was saved, and the
        // components they let through - those held across this load selected, else those a filter typed finds.
        //
        // What is held comes first (code review, 2026-11-04): with a word in the box the list came back with
        // everything it finds selected - a component unticked among them too. Only somebody's choice since lets
        // go of what is held (OnComponentFilterSelectionChanged, ApplyNormalComponentSearchFilter).
        // ###########################################################################################
        private void ShowListsOfBoard(BoardData boardData, string boardKey, string region)
        {
            var categories = ComponentListBuilder.BuildDistinctCategories(boardData);
            this.CategoryFilterListBox.ItemsSource = categories;

            var savedCategories = UserSettings.GetSelectedCategories(boardKey);
            if (savedCategories == null)
            {
                try
                {
                    this.CategoryFilterListBox.SelectAll();
                }
                catch (OutOfMemoryException ex)
                {
                    Logger.Debug(ex, "Failed to apply default category selection - group was too large to select");
                }
            }
            else
            {
                var seenCategories = UserSettings.GetSeenCategories(boardKey);

                for (int i = 0; i < categories.Count; i++)
                {
                    if (SavedCategoryFilter.IsTicked(categories[i], savedCategories, seenCategories))
                    {
                        this.CategoryFilterListBox.Selection.Select(i);
                    }
                }

                // Saved before what was seen was kept: what comes after this shows.
                if (seenCategories == null)
                {
                    UserSettings.SetSeenCategories(boardKey, categories);
                }
            }

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            string searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
            var componentItems = ComponentListBuilder.BuildComponentItems(
                boardData, region, activeCategories, searchTerm, this._currentBoardDraftSummary.DraftedComponentBoardLabels);

            this._componentSearchInEffect = searchTerm;

            bool highlightsWere = this._suppressComponentHighlightUpdate;
            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            this._boardInOtherTabsLists = boardKey;

            if (this._selectionToKeep is { } keep && string.Equals(keep.BoardKey, boardKey, StringComparison.OrdinalIgnoreCase))
            {
                this.SelectInComponentList(componentItems, keep.ComponentKeys);
            }
            else if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try
                {
                    this.ComponentFilterListBox.SelectAll();
                }
                catch
                {
                }
            }

            this._suppressComponentHighlightUpdate = highlightsWere;

            // The rows of the board shown now: they take clicks again (LetGoOfListsForLoad).
            this.OtherTabsListsTakeClicks(true);
        }
    }
}
