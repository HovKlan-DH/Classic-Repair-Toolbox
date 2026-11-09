using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Geometry;
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
    //   - WHAT IS DONE TO THE LISTS HERE STAYS (owner request, 2026-11-03: "if I switch between contribute and
    //     schematics tab, then I do not want it to reset every time ... same for the component selected"): the
    //     categories UNTICKED and the words typed are kept per board for the run (ContributeListState) - unticked
    //     rather than ticked, so a category the draft has just made is ticked. Kept here, they are not the board's
    //     saved filter; handed back to the other tabs' list as the tab is left (below), they are saved as it, as a
    //     tick there is. "Every category ticked the first time" (2026-11-03) is over - see "THE SAME LISTS".
    //   - A COMPONENT PICKED stays picked whatever else the tab shows (TabDrafts.thisPickedComponent) - SEVERAL with Ctrl
    //     or Shift, edited as one (owner request, 2026-11-08; ComponentPane.Several.cs); "Mark all" stays off.
    //   - *** THE SAME LISTS ON EVERY TAB (owner requests, 2026-11-08: "always keep the same selection for the left-side
    //     panel "Category" and "Component" lists", and 2026-11-09: "should keep the exact same components and categories
    //     selected, and it should not change position in the lists ... it should be the exact same lists"). *** Everything
    //     the lists show is handed across as the tab is entered and as it is left, whatever was done there
    //     (TakeListsIntoContribute, GiveListsBackFromContribute): the categories ticked, the word in the filter box, the
    //     components selected - and where each list stands, the same row at its top (ListPosition). So "every category
    //     ticked the first time" (2026-11-03) is over: the tab starts from what the other tabs show.
    //   - THE FILTER BOX ONLY NARROWS THE LIST, selecting nothing - the word handed across narrows the other tabs' list
    //     too, selecting nothing there either; typed there, it selects all it finds, as it always did.
    //   - "WITH ISSUES" COMES WITH THE LIST, there on no other tab (TabDrafts.Problems.cs).
    //   - "Clear" lets go of all of it: every category ticked, nothing typed, nothing narrowed, no component picked.
    //   - THE REGION BUTTONS are the tab's region too (PickRegion) - and follow "Board: Regions" as it is
    //     ticked, so they read what the tab reads.
    //
    // While the tab is shown the other tabs' lists stay behind it; back on either, everything stands where the list shown
    // last stood (owner requests, 2026-11-04: "navigate to the exact same position", and 2026-11-09). What reaches the
    // hidden lists meanwhile is a board read again - what the pane saved - which drops a selection as any load does, so
    // what was selected as the tab was entered, or given back as it was left, is held until such a load has shown it
    // (_selectionToKeep).
    // Its rows stay on screen until the new ones come, and those come selected in one go (LetGoOfListsForLoad,
    // SelectInComponentList; owner report, 2026-11-05).
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // Whether the Contribute tab is shown - so its lists rather than the other tabs'.
        private bool _contributeListsShown;

        // Whether the tab's lists took what the other tabs' show since the tab was entered - not while the board chosen was
        // still being read, when they are taken as its lists come (OnContributeListsShown). Only lists that took them give
        // them back: lists that never did hold nobody's choice.
        private bool _contributeTookLists;

        // True while they are being taken - the tab's lists are read again then, which is no load ending.
        private bool _takingContributeLists;

        // ###########################################################################################
        // What the lists said as they were handed over - the tab's as it was entered, the other tabs' as it was left. A side
        // that still says it was not touched since, and hands nothing over: lists that are the same stay exactly as they
        // were, a pick a word hides on one side included. Null: hand everything over.
        // ###########################################################################################
        private ListsSaid? _contributeSaidAtEntry;
        private ListsSaid? _othersSaidAtExit;

        private sealed record ListsSaid(HashSet<string> Unticked, string Word, HashSet<string> Picked)
        {
            public bool SameAs(ListsSaid other) =>
                this.Unticked.SetEquals(other.Unticked) &&
                string.Equals(this.Word, other.Word, StringComparison.Ordinal) &&
                this.Picked.SetEquals(other.Picked);
        }

        private ListsSaid OtherTabsListsSay()
        {
            var ticked = new HashSet<string>(this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? [], StringComparer.OrdinalIgnoreCase);

            return new ListsSaid(
                new HashSet<string>((this.CategoryFilterListBox.ItemsSource?.Cast<string>() ?? []).Where(category => !ticked.Contains(category)), StringComparer.OrdinalIgnoreCase),
                this.ComponentSearchTextBox.Text ?? string.Empty,
                new HashSet<string>(this.SelectedComponentLabels(), StringComparer.OrdinalIgnoreCase));
        }

        private ListsSaid ContributeListsSay()
        {
            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);

            return new ListsSaid(
                new HashSet<string>(state.Unticked, StringComparer.OrdinalIgnoreCase),
                state.Search,
                new HashSet<string>(this.TabDrafts.PickedForOtherTabs, StringComparer.OrdinalIgnoreCase));
        }

        // Both sets of lists are of the board shown, read: only then is anything handed across.
        private bool ListsAreOfOneBoard =>
            this._currentBoardData is not null &&
            this.ContributeListsShowTheirBoard &&
            string.Equals(this._boardInOtherTabsLists, this._contributeListsBoard, StringComparison.OrdinalIgnoreCase);

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

            // Taken from the other tabs, or changed here, this run: what the lists say on the board when it is shown on this
            // tab again - another board shown here and back takes nothing from the other tabs over it.
            public bool Kept { get; set; }
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
            this._localRegion = RegionChoice.ToShow(BoardRegions.Of(board), UserSettings.Region, UserSettings.RegionOfBoard(boardKey));
            this.UpdateRegionButtonsState();
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
                // A word typed and not taken yet is what the list shows a moment later.
                if (!string.Equals(this.ComponentSearchTextBox.Text ?? string.Empty, this._componentSearchInEffect, StringComparison.Ordinal))
                {
                    this.ApplyComponentSearchNow();
                }

                this.HoldSelectionOfTheOtherTabs();

                // Nothing done on the other tabs since they took this tab's lists: this tab's stay as they were - a pick a
                // word here hides, say, which their list cannot hold.
                if (this._othersSaidAtExit is { } said && said.SameAs(this.OtherTabsListsSay()) && this.ListsAreOfOneBoard)
                {
                    this._contributeTookLists = true;
                }
                else
                {
                    this.TakeListsIntoContribute();
                }

                this._contributeSaidAtEntry = this._contributeTookLists ? this.ContributeListsSay() : null;
            }
            else
            {
                if (!string.Equals(this.ContributeSearchTextBox.Text ?? string.Empty, this._contributeSearchInEffect, StringComparison.Ordinal))
                {
                    this.ApplyContributeSearchNow();
                }

                // Nothing done here since this tab took the other tabs' lists: theirs stay as they were.
                if (this._contributeSaidAtEntry is not { } said || !said.SameAs(this.ContributeListsSay()))
                {
                    this.GiveListsBackFromContribute();
                }

                this._othersSaidAtExit = this._contributeTookLists ? this.OtherTabsListsSay() : null;
            }

            // Where the list shown until now stands, read before it goes - the list shown next is put there.
            (ListBox Components, ListBox Categories) from = shown
                ? (this.ComponentFilterListBox, this.CategoryFilterListBox)
                : (this.ContributeComponentListBox, this.ContributeCategoryListBox);
            (ListBox Components, ListBox Categories) to = shown
                ? (this.ContributeComponentListBox, this.ContributeCategoryListBox)
                : (this.ComponentFilterListBox, this.CategoryFilterListBox);
            ListStand? components = Main.StandOf(from.Components);
            ListStand? categories = Main.StandOf(from.Categories);

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

            // The region one of the board now read - the draft's here, the board on screen's on the other tabs, which can come
            // in other regions (FollowRegionsOfBoardShown).
            this.FollowRegionsOfBoardShown();
            this.UpdateRegionButtonsState();

            if (keyboardInTheBox)
            {
                boxShown.Focus();
            }

            // The lists shown now put where the others stood, once laid out - by the layout pass that comes, never one forced
            // on the whole window here (code review, 2026-11-09: every switch laid the window out on the spot, inside the tab
            // control's selection change).
            this.PutListsWhenLaidOut(to.Components, components, to.Categories, categories);
        }

        // The placement waiting for the next layout pass - let go of as the tab is switched again before it came.
        private EventHandler? _listsWaitingForLayout;

        private ListBox? _listWaitedOn;

        private void PutListsWhenLaidOut(ListBox components, ListStand? componentsStand, ListBox categories, ListStand? categoriesStand)
        {
            if (this._listWaitedOn is { } waited && this._listsWaitingForLayout is { } waiting)
            {
                waited.LayoutUpdated -= waiting;
            }

            this._listsWaitingForLayout = null;
            this._listWaitedOn = null;

            if (componentsStand is null && categoriesStand is null)
            {
                return;
            }

            void Put(object? sender, EventArgs e)
            {
                components.LayoutUpdated -= Put;
                this._listsWaitingForLayout = null;
                this._listWaitedOn = null;

                Main.PutAt(components, componentsStand);
                Main.PutAt(categories, categoriesStand);
            }

            this._listsWaitingForLayout = Put;
            this._listWaitedOn = components;
            components.LayoutUpdated += Put;
        }

        // ---- one set of lists, as far as anybody can tell -----------------------------------------------

        // ###########################################################################################
        // *** THE SAME LISTS ON EVERY TAB (owner request, 2026-11-09: "Toggling between tabs (e.g. "Contribute" and
        // "Schematics") should keep the exact same components and categories selected, and it should not change position
        // in the lists. Right now it works independently, having each its component scope and lists, though it should be
        // the exact same lists"). *** As the tab is entered its lists take everything the other tabs' show: the categories
        // ticked, the word in the filter box - which narrows here - and the components selected, picked here. Not while
        // the board chosen is still being read: the lists then are of the board before (ContributeListsShowTheirBoard) -
        // they are taken as its lists come (OnContributeListsShown).
        // ###########################################################################################
        private void TakeListsIntoContribute()
        {
            this._contributeTookLists = false;

            if (!this.ListsAreOfOneBoard)
            {
                return;
            }

            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);
            var ticked = new HashSet<string>(this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? [], StringComparer.OrdinalIgnoreCase);

            state.Unticked.Clear();
            state.Unticked.UnionWith((this.CategoryFilterListBox.ItemsSource?.Cast<string>() ?? []).Where(category => !ticked.Contains(category)));
            state.Search = this.ComponentSearchTextBox.Text ?? string.Empty;
            state.Kept = true;

            this._takingContributeLists = true;

            try
            {
                this._contributeSearchTimer?.Stop();
                this.ShowContributeSearchText(state.Search);
                this.TabDrafts.UseComponentFilter(state.Unticked, state.Search);
                this.ShowContributeTicks();
                this.TabDrafts.TakePickFromOtherTabs(this.SelectedComponentLabels(), everyRowOfTheirs: this.IsEveryComponentRowSelected);
            }
            finally
            {
                this._takingContributeLists = false;
            }

            this._contributeTookLists = true;
        }

        // ###########################################################################################
        // ... and as it is left, the other tabs' take the tab's: the categories ticked (saved as the board's filter, as a
        // tick there is), the word - narrowing their list as it narrows this one, never selecting all it finds, as typing it
        // there does - and the components picked here, selected and lit up on the picture. Held for the board read again
        // as the tab is left, as what was selected as it was entered was (_selectionToKeep).
        // ###########################################################################################
        private void GiveListsBackFromContribute()
        {
            if (!this._contributeTookLists || !this.ListsAreOfOneBoard || this._currentBoardData is not { } board)
            {
                return;
            }

            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);
            List<string> categories = this.CategoryFilterListBox.ItemsSource?.Cast<string>().ToList() ?? [];
            List<string> ticked = categories.Where(category => !state.Unticked.Contains(category)).ToList();
            var selection = this.CategoryFilterListBox.Selection;
            bool ticksChange = categories.Where((category, index) => selection.IsSelected(index) == state.Unticked.Contains(category)).Any();
            bool wordChanges = !string.Equals(this.ComponentSearchTextBox.Text ?? string.Empty, state.Search, StringComparison.Ordinal);
            string boardKey = this.GetCurrentBoardKey();

            if (ticksChange)
            {
                this._suppressCategoryFilterSave = true;
                selection.BeginBatchUpdate();

                try
                {
                    for (int index = 0; index < categories.Count; index++)
                    {
                        if (state.Unticked.Contains(categories[index]))
                        {
                            selection.Deselect(index);
                        }
                        else
                        {
                            selection.Select(index);
                        }
                    }
                }
                finally
                {
                    selection.EndBatchUpdate();
                    this._suppressCategoryFilterSave = false;
                }

                if (boardKey.Length > 0)
                {
                    UserSettings.SetSelectedCategories(boardKey, ticked, categories);
                }
            }

            if (wordChanges)
            {
                this._componentSearchDebounceTimer?.Stop();
                this.ShowComponentSearchText(state.Search);
            }

            var wanted = new HashSet<string>(this.TabDrafts.PickedForOtherTabs, StringComparer.OrdinalIgnoreCase);

            this._suppressComponentHighlightUpdate = true;

            try
            {
                if (ticksChange || wordChanges || this.ComponentFilterListBox.ItemsSource is null)
                {
                    this.ComponentFilterListBox.ItemsSource = ComponentListBuilder.BuildComponentItems(
                        board,
                        this._localRegion,
                        new HashSet<string>(ticked, StringComparer.OrdinalIgnoreCase),
                        state.Search,
                        this._currentBoardDraftSummary.DraftedComponentBoardLabels);
                }

                List<ComponentListItem> rows = this.ComponentFilterListBox.ItemsSource?.Cast<ComponentListItem>().ToList() ?? [];
                var components = this.ComponentFilterListBox.Selection;
                components.BeginBatchUpdate();

                try
                {
                    components.Clear();

                    for (int index = 0; index < rows.Count; index++)
                    {
                        if (wanted.Contains(rows[index].BoardLabel))
                        {
                            components.Select(index);
                        }
                    }
                }
                finally
                {
                    components.EndBatchUpdate();
                }
            }
            finally
            {
                this._suppressComponentHighlightUpdate = false;
            }

            var keys = new HashSet<string>(
                this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>().Select(item => item.SelectionKey) ?? [],
                StringComparer.OrdinalIgnoreCase);

            this._selectionToKeep = keys.Count > 0 || state.Search.Length > 0 ? (boardKey, keys) : null;

            this.TabSchematicsControl.UpdateHighlightsForComponents(this.SelectedComponentLabels());
            this.TabOverview.ApplyFilter(state.Search);
        }

        // Where a list stands: its rows' keys, the row at its top and how far into it (ListPosition).
        private sealed record ListStand(IReadOnlyList<string> Keys, int Row, double Into);

        private static IReadOnlyList<string> KeysOf(ListBox list) =>
            (list.ItemsSource?.Cast<object>() ?? [])
                .Select(item => item switch
                {
                    ComponentListItem component => component.SelectionKey,
                    string category => category,
                    _ => item?.ToString() ?? string.Empty,
                })
                .ToList();

        private static ScrollViewer? ScrollerOf(ListBox list) => list.GetVisualDescendants().OfType<ScrollViewer>().FirstOrDefault();

        // Null for a list never laid out, or with nothing to measure.
        private static ListStand? StandOf(ListBox list)
        {
            IReadOnlyList<string> keys = Main.KeysOf(list);

            return Main.ScrollerOf(list) is { } scroller &&
                   ListPosition.TopOf(scroller.Offset.Y, scroller.Extent.Height, keys.Count) is { } top
                ? new ListStand(keys, top.Row, top.Into)
                : null;
        }

        // `list` put where `stand` was: the same row at its top, as far into it.
        private static void PutAt(ListBox list, ListStand? stand)
        {
            if (stand is null || Main.ScrollerOf(list) is not { } scroller)
            {
                return;
            }

            IReadOnlyList<string> keys = Main.KeysOf(list);
            int row = ListPosition.MatchingRow(stand.Keys, stand.Row, keys);

            if (row >= 0)
            {
                scroller.Offset = new Avalonia.Vector(
                    scroller.Offset.X,
                    ListPosition.OffsetOf(row, stand.Into, scroller.Extent.Height, keys.Count, scroller.Viewport.Height));
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

        // Whether the other tabs' component list holds "Mark all"'s selection - the button pressed and every row, more than one,
        // still selected - which the Contribute tab shows and never opens (TabDrafts.TakePickFromOtherTabs).
        private bool IsEveryComponentRowSelected =>
            this._markAllPressed &&
            this.ComponentFilterListBox.ItemCount > 1 &&
            (this.ComponentFilterListBox.SelectedItems?.Count ?? 0) == this.ComponentFilterListBox.ItemCount;

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
            state.Kept = true;

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
            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);

            this._contributeSearchInEffect = word;

            if (!string.Equals(state.Search, word, StringComparison.Ordinal))
            {
                state.Search = word;
                state.Kept = true;
            }
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

            // Another board's lists are the other tabs' to give, as they come (OnContributeListsShown) - never this
            // board's handed to them as the tab is left.
            this._contributeTookLists = false;
            this._contributeSaidAtEntry = null;
            this._othersSaidAtExit = null;
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

        // ###########################################################################################
        // The lists were read again: the region buttons read what the tab reads ("Board: Regions" ticked shows them at
        // once), and the ticks follow the categories there now are. The tab entered while the board chosen was still being
        // read takes the other tabs' lists now that its own are of that board (TakeListsIntoContribute).
        // ###########################################################################################
        private void OnContributeListsShown(object? sender, EventArgs e)
        {
            // A region the board no longer has is let go of first - which reads the lists again, and runs this then.
            if (this.FollowRegionsOfBoardShown())
            {
                return;
            }

            this.UpdateRegionButtonsState();
            this.ShowContributeTicks();

            if (this._contributeListsShown && !this._contributeTookLists && !this._takingContributeLists && this.ListsAreOfOneBoard)
            {
                // A board shown here before this run keeps what its lists said here - given back whole as the tab is left.
                if (this.ContributeListStateOf(this._contributeListsBoard).Kept)
                {
                    this._contributeTookLists = true;
                    return;
                }

                this.TakeListsIntoContribute();
                this._contributeSaidAtEntry = this._contributeTookLists ? this.ContributeListsSay() : null;
            }
        }

        // "Clear" on the tab: back to how it starts - nothing narrowed, and no component picked.
        private void ClearContributeLists()
        {
            this._contributeSearchTimer?.Stop();

            ContributeListState state = this.ContributeListStateOf(this._contributeListsBoard);

            state.Unticked.Clear();
            state.Search = string.Empty;
            state.Kept = true;

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
            // A key no row has any more - its component's names changed on the Contribute tab since it was held (code review,
            // 2026-11-09: as the board was read again it lost its selection, and coming back to the tab closed its pane) - is
            // found by its label and region, the row it was.
            var present = new HashSet<string>(rows.Select(row => row.SelectionKey), StringComparer.OrdinalIgnoreCase);
            var renamed = new HashSet<string>(
                selectionKeys.Where(key => !present.Contains(key)).Select(Main.LabelAndRegionOf),
                StringComparer.OrdinalIgnoreCase);

            this.ComponentFilterListBox.Selection.BeginBatchUpdate();

            try
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    if (selectionKeys.Contains(rows[i].SelectionKey) ||
                        (renamed.Count > 0 && renamed.Contains(Main.LabelAndRegionOf(rows[i].SelectionKey))))
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

        // A component list row's key (ComponentListItem.SelectionKey: label, friendly name, technical name, region) as its
        // label and region alone - what stays as its names change.
        internal static string LabelAndRegionOf(string selectionKey)
        {
            string[] parts = selectionKey.Split('\u001F');

            return parts.Length >= 2 ? string.Concat(parts[0], "\u001F", parts[^1]) : selectionKey;
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
