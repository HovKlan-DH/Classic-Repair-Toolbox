using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // "BOARD: REGIONS" AND "BOARD: CATEGORIES" (owner requests, 2026-11-01), two of the board's own parts.
    //
    //   - REGIONS: the one tick "This board comes with multiple regions" and its regions listed, in their
    //     colors and in the buttons' order, dragged up or down (owner request, 2026-11-09; BoardRegions - stored
    //     in the board file and published). The tick cannot be taken away while a row says a region: such a row
    //     would go on being there with no button left to reach it.
    //   - CATEGORIES: rename, add and delete the words the board's components are sorted under
    //     (BoardCategoryEdits). No list is stored: a category is what its components carry, so a rename
    //     or a delete is one step changing those rows, and one no component uses yet is held here, by
    //     the board, until something does or CRT closes (owner decision: "kept until first used").
    //     A name commits on Enter or when its box is left, never on a pause in typing.
    //   - AND THEIR ORDER (owner request, 2026-11-01: "rearrange the category up/down, alike Files or
    //     Links"): dragged by a handle or Alt+Up/Down, one step moving its components' rows - the order
    //     IS which is seen first (BoardCategoryEdits.InOrder). One nothing uses has no rows to move, so
    //     no handle, and the components with no category stay on top, as what needs doing.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        private readonly ObservableCollection<ComponentPaneCategoryRow> thisCategories = new();

        // The categories added and not used yet, per board - for the run. By the board's identity, not
        // its session: a draft made only to look goes again, and they outlive it.
        private readonly Dictionary<string, List<string>> thisHeldCategories = new(StringComparer.OrdinalIgnoreCase);

        internal IReadOnlyList<ComponentPaneCategoryRow> CategoriesForTests => this.thisCategories;

        // A category's components now under another - the host's, whose filter may have to follow.
        internal event EventHandler<(string From, string To)>? CategoryMoved;

        // ---- Board: Regions --------------------------------------------------------------------------

        private readonly ObservableCollection<ComponentPaneRegionRow> thisRegionRows = new();

        // ###########################################################################################
        // ONE TICK, "This board comes with multiple regions", AND EVERY REGION LISTED WITH ITS COLORS (owner request,
        // 2026-11-09: "Why not make this a little differently, and then have the checkbox state "This board comes with
        // multiple regions" and then list all regions and its colors? It should be possible to reorder the list of regions
        // up/down" - a PAL and NTSC tick with the board's own regions under it until then): each region as its button looks
        // under the component list (RegionChoice.ColorsOf - PAL's and NTSC's own colors, the others taking four in turn by
        // their place among them), how many rows hold it (BoardRegions.RowsHolding - a component told apart in every region
        // goes with it), "Delete region" while none does and more than two are left, and a
        // handle to drag it up or down - the buttons' order. Added by name under the list. Unticked, the board comes in
        // none, which a row saying a region keeps from happening.
        // ###########################################################################################
        private void ShowRegionsPart(BoardData board)
        {
            IReadOnlyList<string> regions = BoardRegions.Of(board);
            int rows = BoardRegions.RegionalRows(board);
            bool has = regions.Count > 0;

            this.RegionsCheckBox.Content = ComponentPaneWording.RegionsHasOnBoard;
            this.RegionsCheckBox.IsChecked = has;
            this.RegionsCheckBox.IsEnabled = !has || rows == 0;
            BoldRuns.ShowMarked(this.RegionsInUseText, has && rows > 0 ? ComponentPaneWording.RegionsInUse(rows) : string.Empty);
            this.RegionsInUseText.IsVisible = has && rows > 0;
            this.RegionsListPanel.IsVisible = has;
            this.AddRegionPanel.IsVisible = has && regions.Count < BoardRegions.MostRegions;

            // Every region's facts worked out in one go (BoardRegions.HoldingOf) - and its colors: the board's own where it gives
            // any (RegionColors, owner request 2026-11-11), else the ones CRT gives it, which the pickers start from.
            List<ComponentPaneRegionRow.Facts> wanted = BoardRegions.HoldingOf(board)
                .Select(region =>
                {
                    string colors = RegionChoice.ColorsOf(regions, region.Region);
                    RegionColorEntry? own = RegionColors.RowOf(board, region.Region);

                    return new ComponentPaneRegionRow.Facts(region.Region, region.Rows, region.CanDelete, colors, region.Going)
                    {
                        Components = region.Components,
                        BoardFiles = region.BoardFiles,
                        BoardLinks = region.BoardLinks,
                        OwnBackground = RegionColors.Normalised(own?.Background),
                        OwnText = RegionColors.Normalised(own?.Background) is null ? null : RegionColors.Normalised(own?.Text),
                        DefaultBackground = this.ColorOfResource($"{colors}_Bg"),
                        DefaultText = this.ColorOfResource($"{colors}_Fg"),
                    };
                })
                .ToList();

            this.thisShowingRegionRows = true;

            try
            {
                InPlaceList.Sync(this.thisRegionRows, wanted, (row, facts) => row.Shows(facts), this.NewRegionRow);

                for (int index = 0; index < wanted.Count && index < this.thisRegionRows.Count; index++)
                {
                    this.thisRegionRows[index].Use(wanted[index]);
                }
            }
            finally
            {
                this.thisShowingRegionRows = false;
            }
        }

        // True while the rows are given what the board holds - their colors changing then is nothing picked.
        private bool thisShowingRegionRows;

        // A row made for a region, its colors picked followed (OnRegionRowChanged).
        private ComponentPaneRegionRow NewRegionRow(ComponentPaneRegionRow.Facts facts)
        {
            var row = new ComponentPaneRegionRow(facts);
            row.PropertyChanged += this.OnRegionRowChanged;

            return row;
        }

        // The color a theme resource gives - for a region's pickers to start from. Gray for none.
        private Avalonia.Media.Color ColorOfResource(string key) =>
            this.TryFindResource(key, this.ActualThemeVariant, out object? value) && value is Avalonia.Media.ISolidColorBrush brush
                ? brush.Color
                : Avalonia.Media.Colors.Gray;

        // ###########################################################################################
        // A REGION'S COLORS PICKED (owner request, 2026-11-11: "In "Board: Regions" I would like to be able to set the background
        // color, and then you should calculate the complementary text color ... However, I would also like to be able to set the
        // text color"): into the draft once picking pauses, as a picture's colors go (the typing timer, CommitRegionColors) - a
        // picker dragged over its spectrum changes the color many times a second.
        // ###########################################################################################
        private void OnRegionRowChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (this.thisShowing || this.thisShowingRegionRows || !ComponentPaneRegionRow.IsColorChange(e.PropertyName))
            {
                return;
            }

            this.thisTypingTimer.Stop();
            this.thisTypingTimer.Start();
        }

        // Every region's colors as their rows hold them, into the draft as ONE step - none when nothing changed.
        internal void CommitRegionColors()
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            List<(string Name, string? Background, string? Text)> wanted = this.thisRegionRows
                .Select(row => (row.Name, row.OwnBackground, row.OwnText))
                .ToList();

            BoardData Edit(BoardData board) =>
                wanted.Aggregate(board, (current, region) => RegionColors.WithColors(current, region.Name, region.Background, region.Text));

            if (ReferenceEquals(Edit(session.Content.Board), session.Content.Board))
            {
                return;
            }

            this.ChangeRegions(Edit, DraftStepWording.RegionColors);
        }

        // "Automatic text color": the region's text in the color that reads best on its background - one step.
        private void OnAutomaticTextColorClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneRegionRow row })
            {
                this.thisTypingTimer.Stop();
                row.UseAutomaticText();
                this.CommitRegionColors();
            }
        }

        // "Reset colors": the region back in the colors CRT gives it - one step.
        private void OnResetRegionColorsClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneRegionRow row })
            {
                this.thisTypingTimer.Stop();
                row.UseCrtColors();
                this.CommitRegionColors();
            }
        }

        // "Show all" and "Show fewer" under a region's components.
        private void OnRegionShowAllClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneRegionRow row })
            {
                row.ShowAll = !row.ShowAll;
            }
        }

        // A component a region is applied on, opened - in that region, so its rows of the region are the ones shown.
        internal event EventHandler<(string Label, string Region)>? OpenComponentRequested;

        private void OnRegionComponentClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneRegionComponent component })
            {
                this.CommitHeld();
                this.OpenComponentRequested?.Invoke(this, (component.Label, component.Region));
            }
        }

        internal IReadOnlyList<ComponentPaneRegionRow> RegionRowsForTests => this.thisRegionRows;

        private void OnRegionsCheckClick(object? sender, RoutedEventArgs e) => this.SetRegions(this.RegionsCheckBox.IsChecked == true);

        internal void SetRegions(bool has) => this.ChangeRegions(board => BoardRegions.WithRegions(board, has));

        private void ChangeRegions(Func<BoardData, BoardData> edit, string step = DraftStepWording.Regions)
        {
            if (this.thisSession is { } session)
            {
                // A color picked and not in the draft yet (the typing pause) goes in first, as its own step: the rows are
                // given the board's colors again below, which would take it back without a word.
                if (step != DraftStepWording.RegionColors)
                {
                    this.CommitRegionColors();
                }

                this.Apply(step, edit);
                this.ShowRegionsPart(session.Content.Board);
                this.ShowBoardPartChanged();
            }
        }

        // A row's handle pressed - the drag starts once the pointer moves (ListRowDrag.Press).
        private void OnRegionHandlePressed(object? sender, PointerPressedEventArgs e) => this.thisRegionsDrag?.Press(sender, e);

        // A region moved one place up (-1) or down (+1) - Alt+Up and Alt+Down on its row; false when it cannot.
        internal bool MoveRegion(ComponentPaneRegionRow row, int by)
        {
            ArgumentNullException.ThrowIfNull(row);

            int index = this.thisRegionRows.IndexOf(row);
            int to = index + by;

            if (index < 0 || to < 0 || to >= this.thisRegionRows.Count)
            {
                return false;
            }

            this.thisRegionRows.Move(index, to);
            this.CommitRegionOrder();

            return true;
        }

        // The regions as they now stand on screen, into the draft as one step - the buttons in that order.
        private void CommitRegionOrder()
        {
            List<string> order = this.thisRegionRows.Select(row => row.Name).ToList();

            this.ChangeRegions(board => BoardRegions.InOrder(board, order));
        }

        // ###########################################################################################
        // A region of the board's own added - its name by BoardRegions.NameProblem's rule, the one the server holds a
        // submission to; refused, the name stays in its box with the reason under it.
        // ###########################################################################################
        internal void AddRegion(string? name)
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            string? problem = BoardRegions.NameProblem(name, BoardRegions.Of(session.Content.Board));

            this.NewRegionProblemText.Text = problem ?? string.Empty;
            this.NewRegionProblemText.IsVisible = problem is not null;

            if (problem is not null)
            {
                return;
            }

            this.ChangeRegions(board => BoardRegions.WithRegion(board, name));
            this.NewRegionNameBox.Text = string.Empty;
        }

        private void OnAddRegionClick(object? sender, RoutedEventArgs e) => this.AddRegion(this.NewRegionNameBox.Text);

        private void OnNewRegionNameKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                this.AddRegion(this.NewRegionNameBox.Text);
            }
        }

        private void OnDeleteRegionClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneRegionRow row } && row.CanDelete)
            {
                this.DeleteRegion(row.Name);
            }
        }

        internal void DeleteRegion(string region) => this.ChangeRegions(board => BoardRegions.WithoutRegion(board, region));

        // ---- Board: Categories -----------------------------------------------------------------------

        // The categories added and not used yet of the board open - none any component has.
        private List<string> HeldCategories(BoardData board)
        {
            if (this.thisSession is not { } session ||
                !this.thisHeldCategories.TryGetValue(session.ExcelDataFile, out List<string>? held))
            {
                return [];
            }

            held.RemoveAll(name => board.Components.Any(row => string.Equals(row.Category?.Trim(), name, StringComparison.OrdinalIgnoreCase)));

            return held;
        }

        // A new board renamed: those waiting for their first component go with it, under its new identity.
        internal void FollowBoardRenamed(string from, string to)
        {
            if (this.thisHeldCategories.Remove(from, out List<string>? held))
            {
                this.thisHeldCategories[to] = held;
            }
        }

        // What a category box offers: the board's, and those added and waiting for their first component.
        private List<string> CategoryChoices(BoardData board) =>
            ComponentListBuilder.BuildDistinctCategories(board)
                .Concat(this.HeldCategories(board))
                .OrderBy(category => category, StringComparer.OrdinalIgnoreCase)
                .ToList();

        // ###########################################################################################
        // The same with `own` among them as it is spelt, the word a component's row holds: a list only
        // able to select the board's first spelling would read "ic" as a change to "IC" nobody made.
        // ###########################################################################################
        private List<string> CategoryChoicesWith(BoardData board, string? own)
        {
            List<string> choices = this.CategoryChoices(board);

            if (!string.IsNullOrEmpty(own) && !choices.Contains(own, StringComparer.Ordinal))
            {
                int place = choices.FindIndex(choice => string.Compare(choice, own, StringComparison.OrdinalIgnoreCase) > 0);
                choices.Insert(place < 0 ? choices.Count : place, own);
            }

            return choices;
        }

        // ###########################################################################################
        // The rows, changed in place: a name commits as its box loses focus - on the PRESS of another
        // row's "Delete" - and a row made again would throw that click away.
        // ###########################################################################################
        private void ShowCategoriesPart(BoardData board)
        {
            List<BoardCategoryRow> wanted = BoardCategoryEdits.Of(board)
                .Concat(this.HeldCategories(board).Select(name => new BoardCategoryRow(name, 0)))
                .ToList();

            InPlaceList.Sync(
                this.thisCategories,
                wanted,
                (row, category) => row.Shows(category),
                category => new ComponentPaneCategoryRow(category));

            List<string> names = wanted.Select(row => row.Name).Where(name => name.Length > 0).ToList();

            foreach (ComponentPaneCategoryRow row in this.thisCategories)
            {
                row.UseOtherCategories(names);
            }
        }

        // A row's handle pressed - the drag starts once the pointer moves (ListRowDrag.Press).
        private void OnCategoryHandlePressed(object? sender, PointerPressedEventArgs e) => this.thisCategoriesDrag?.Press(sender, e);

        // ###########################################################################################
        // A category moved one place up (-1) or down (+1), past the next one that can move - false when
        // it cannot. The rows are moved in place, so the list keeps its boxes.
        // ###########################################################################################
        internal bool MoveCategory(ComponentPaneCategoryRow row, int by)
        {
            ArgumentNullException.ThrowIfNull(row);

            this.CommitHeld();

            int index = this.thisCategories.IndexOf(row);
            int to = index + by;

            if (index < 0 || !row.CanMove || to < 0 || to >= this.thisCategories.Count || !this.thisCategories[to].CanMove)
            {
                return false;
            }

            this.thisCategories.Move(index, to);
            this.CommitCategoryOrder();

            return true;
        }

        // ###########################################################################################
        // The categories as they now stand on screen, into the draft as one step. A row dropped among
        // those that cannot move goes back beside the ones that can as the part is read again.
        // ###########################################################################################
        private void CommitCategoryOrder()
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            List<string> names = this.thisCategories.Where(row => row.CanMove).Select(row => row.Name).ToList();

            this.Apply(DraftStepWording.CategoriesOrdered, board => BoardCategoryEdits.InOrder(board, names));
            this.ShowCategoriesPart(session.Content.Board);
            this.ShowBoardPartChanged();
        }

        private void OnCategoryNameKeyDown(object? sender, KeyEventArgs e)
        {
            if (sender is not Control { DataContext: ComponentPaneCategoryRow row })
            {
                return;
            }

            if (e.Key == Key.Enter)
            {
                this.CommitCategoryName(row);
            }
            else if (e.Key == Key.Escape)
            {
                row.PutNameBack();
            }
        }

        private void OnCategoryNameLostFocus(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { DataContext: ComponentPaneCategoryRow row })
            {
                this.CommitCategoryName(row);
            }
        }

        // ###########################################################################################
        // A category's name as typed: every component under it moves with it, as one step. One no
        // component has only changes the name held. A blank name is refused, the reason under it.
        // ###########################################################################################
        internal void CommitCategoryName(ComponentPaneCategoryRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (this.thisSession is not { } session || string.Equals(row.NameDraft.Trim(), row.Name, StringComparison.Ordinal))
            {
                row.Problem = string.Empty;
                return;
            }

            string typed = row.NameDraft.Trim();

            if (typed.Length == 0)
            {
                row.Problem = ComponentPaneWording.CategoryNameEmpty;
                return;
            }

            row.Problem = string.Empty;

            BoardData board = session.Content.Board;

            if (row.Components == 0)
            {
                List<string> held = this.HeldCategories(board);
                held.RemoveAll(name => string.Equals(name, row.Name, StringComparison.OrdinalIgnoreCase));

                if (!board.Components.Any(component => string.Equals(component.Category?.Trim(), typed, StringComparison.OrdinalIgnoreCase)))
                {
                    this.HoldCategory(session, typed);
                }

                row.UseName(typed);
                this.ShowCategoriesPart(session.Content.Board);
                return;
            }

            string from = row.Name;
            string to = BoardCategoryEdits.Spelling(board, from, typed);

            // Named before the step: the rows are matched by what they show as the part is read again.
            row.UseName(to);

            this.Apply(DraftStepWording.CategoryRenamed(from, to), current => BoardCategoryEdits.Renamed(current, from, to));
            this.CategoryMoved?.Invoke(this, (from, to));
            this.ShowCategoriesPart(session.Content.Board);
            this.ShowBoardPartChanged();
        }

        // ###########################################################################################
        // Every name still in its box, as its box's own Enter or leave takes it (code review, 2026-11-02:
        // closing CRT by the title bar moves no focus, and CommitHeld passed this part by - nothing
        // unsaved, nothing asked, the name gone). A blank one stays refused, the reason under it. Each row
        // asked as it now stands: a commit changes the rows in place, and one joining another goes.
        // ###########################################################################################
        private void CommitCategoryNames()
        {
            foreach (ComponentPaneCategoryRow row in this.thisCategories.ToList())
            {
                if (this.thisCategories.Contains(row) && !string.Equals(row.NameDraft.Trim(), row.Name, StringComparison.Ordinal))
                {
                    this.CommitCategoryName(row);
                }
            }
        }

        private void HoldCategory(DraftSession session, string name)
        {
            if (!this.thisHeldCategories.TryGetValue(session.ExcelDataFile, out List<string>? held))
            {
                held = new List<string>();
                this.thisHeldCategories[session.ExcelDataFile] = held;
            }

            // Named as one held already, it joins it: two rows of one name offered it twice, and deleting
            // one took both.
            if (!held.Contains(name, StringComparer.OrdinalIgnoreCase))
            {
                held.Add(name);
            }
        }

        private void OnDeleteCategoryClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneCategoryRow row })
            {
                this.AskDeleteCategory(row);
            }
        }

        // One nothing uses goes at once; one with components asks where they go, as none may be left
        // without a category.
        internal void AskDeleteCategory(ComponentPaneCategoryRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (this.thisSession is not { } session)
            {
                return;
            }

            if (row.Components > 0)
            {
                row.IsDeleting = true;
                return;
            }

            this.HeldCategories(session.Content.Board).RemoveAll(name => string.Equals(name, row.Name, StringComparison.OrdinalIgnoreCase));
            this.ShowCategoriesPart(session.Content.Board);
        }

        private void OnCancelDeleteCategoryClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneCategoryRow row })
            {
                row.IsDeleting = false;
            }
        }

        private void OnConfirmDeleteCategoryClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneCategoryRow row })
            {
                this.DeleteCategory(row);
            }
        }

        // The category gone, its components under the one picked - one step, which "Undo" takes back.
        internal void DeleteCategory(ComponentPaneCategoryRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (this.thisSession is not { } session || row.MoveTo is not { Length: > 0 } to)
            {
                return;
            }

            string from = row.Name;

            this.Apply(DraftStepWording.CategoryDeleted(from, to), current => BoardCategoryEdits.Renamed(current, from, to));
            this.CategoryMoved?.Invoke(this, (from, to));
            this.ShowCategoriesPart(session.Content.Board);
            this.ShowBoardPartChanged();
        }

        private void OnAddCategoryClick(object? sender, RoutedEventArgs e) => this.AddCategory();

        private void OnNewCategoryKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.AddCategory();
                e.Handled = true;
            }
        }

        // A new category, held until a component uses it. Its name is never one the board has.
        internal void AddCategory()
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            string typed = this.NewCategoryNameBox.Text?.Trim() ?? string.Empty;
            string? problem = typed.Length == 0
                ? ComponentPaneWording.CategoryNameEmpty
                : this.thisCategories.Any(row => string.Equals(row.Name, typed, StringComparison.OrdinalIgnoreCase))
                    ? ComponentPaneWording.CategoryAlreadyThere(typed)
                    : null;

            this.NewCategoryProblemText.Text = problem ?? string.Empty;
            this.NewCategoryProblemText.IsVisible = problem is not null;

            if (problem is not null)
            {
                return;
            }

            this.HoldCategory(session, typed);
            this.NewCategoryNameBox.Text = string.Empty;
            this.ShowCategoriesPart(session.Content.Board);
            this.ShowBoardPartTitle();
        }
    }

    // ###########################################################################################
    // One category as "Board: Categories" holds it: the name in its box, how many components carry it,
    // and - while it is being deleted - where they go. A blank name is the components with none.
    // ###########################################################################################
    public sealed class ComponentPaneCategoryRow : ComponentPaneDraggableRow
    {
        private string thisName;
        private string thisNameDraft;
        private string thisProblem = string.Empty;
        private bool thisIsDeleting;
        private string? thisMoveTo;
        private IReadOnlyList<string> thisMoveChoices = [];

        public ComponentPaneCategoryRow(BoardCategoryRow category)
        {
            this.thisName = category.Name;
            this.thisNameDraft = category.Name;
            this.Components = category.Components;
        }

        public int Components { get; }

        public string Name => this.thisName;

        // Never null, whatever a box bound to it writes: every commit of what the pane holds reads it now.
        public string NameDraft { get => this.thisNameDraft; set => this.Set(ref this.thisNameDraft, value ?? string.Empty); }

        public string CountText => this.Components > 0 ? ComponentPaneWording.CategoryCount(this.Components) : ComponentPaneWording.CategoryUnused;

        // The components with no category: an error until they have one, and nothing to delete.
        public bool NeedsOne => this.thisName.Length == 0;

        public bool CanDelete => !this.NeedsOne;

        // Only a category components carry has rows to move.
        public bool CanMove => !this.NeedsOne && this.Components > 0;

        public string Problem
        {
            get => this.thisProblem.Length > 0 ? this.thisProblem : this.NeedsOne ? ComponentPaneWording.CategoryRowsNeedOne : string.Empty;
            set
            {
                this.thisProblem = value ?? string.Empty;
                this.Raise(nameof(this.Problem));
                this.Raise(nameof(this.HasProblem));
            }
        }

        public bool HasProblem => this.Problem.Length > 0;

        public bool IsDeleting
        {
            get => this.thisIsDeleting;
            set
            {
                this.thisIsDeleting = value;
                this.Raise(nameof(this.IsDeleting));
            }
        }

        public IReadOnlyList<string> MoveChoices => this.thisMoveChoices;

        public bool HasMoveChoices => this.thisMoveChoices.Count > 0;

        public string? MoveTo
        {
            get => this.thisMoveTo;
            set
            {
                if (value is not null)
                {
                    this.thisMoveTo = value;
                    this.Raise(nameof(this.MoveTo));
                    this.Raise(nameof(this.CanDeleteNow));
                }
            }
        }

        public bool CanDeleteNow => !string.IsNullOrEmpty(this.thisMoveTo);

        public string DeleteQuestion => this.HasMoveChoices
            ? ComponentPaneWording.CategoryDeleteAsk(this.Components)
            : ComponentPaneWording.CategoryDeleteNeedsAnother;

        // Whether this row already shows `category`, so it need not be made again.
        public bool Shows(BoardCategoryRow category) =>
            string.Equals(this.thisName, category.Name, StringComparison.Ordinal) && this.Components == category.Components;

        // The name as it now is - what the box holds, once taken.
        public void UseName(string name)
        {
            this.thisName = name;
            this.thisNameDraft = name;
            this.Raise(nameof(this.Name));
            this.Raise(nameof(this.NameDraft));
            this.Raise(nameof(this.NeedsOne));
            this.Raise(nameof(this.CanDelete));
            this.Raise(nameof(this.CanMove));
            this.Raise(nameof(this.Problem));
            this.Raise(nameof(this.HasProblem));
        }

        public void PutNameBack()
        {
            this.NameDraft = this.thisName;
            this.Problem = string.Empty;
        }

        // Where its components could go: every other category of the board - in place, as a list
        // replaced lets go of what is picked in it.
        public void UseOtherCategories(IEnumerable<string> names)
        {
            List<string> others = names.Where(name => !string.Equals(name, this.thisName, StringComparison.OrdinalIgnoreCase)).ToList();

            if (others.SequenceEqual(this.thisMoveChoices, StringComparer.Ordinal))
            {
                return;
            }

            this.thisMoveChoices = others;
            this.Raise(nameof(this.MoveChoices));
            this.Raise(nameof(this.HasMoveChoices));
            this.Raise(nameof(this.DeleteQuestion));

            if (this.thisMoveTo is not null && !others.Contains(this.thisMoveTo, StringComparer.Ordinal))
            {
                this.thisMoveTo = null;
                this.Raise(nameof(this.MoveTo));
                this.Raise(nameof(this.CanDeleteNow));
            }
        }
    }

    // ###########################################################################################
    // A region on "Board: Regions" (owner request, 2026-11-09): its name in its button's colors and whether it can be
    // deleted - why not said under it, how many rows have it applied among it. Dragged up or down by its handle. The colors are the markup's, by a
    // class per set (Border.RegionPill), so they follow the theme and what a person set for PAL and NTSC as the buttons do.
    // ###########################################################################################
    public sealed class ComponentPaneRegionRow : ComponentPaneDraggableRow
    {
        // How many components it is applied on are named before "Show all" (owner request, 2026-11-11: "Can you show those
        // components (all of them) if it is less than X components, and if it has more then it shows something like "... and
        // <b>X</b> more - click to see all."" - 10, then 7 the same day: "maybe have 7 instead of 10").
        public const int ComponentsNamedAtMost = 7;

        // The height of a region's line: its pill and its two color pickers alike (owner request, 2026-11-11) - a picked region
        // button's (12pt text, 4px padding, a 2px edge).
        public const double LineHeight = 28;

        // ###########################################################################################
        // What a row shows. The same NAME, the same row (InPlaceList), its facts changed in place (Use) - so a drag's row, and a
        // row the keyboard is in, survive the step they make (code review, 2026-11-09: moved, two regions swapped colors, their
        // rows were made again, and the focus Alt+Up had kept was lost).
        //
        // Since 2026-11-11 also: the components its rows are of, and the board files and links with it applied (BoardRegions
        // .RegionHolding); the colors the board gives it - "#RRGGBB", null for none - and the ones CRT gives it, which the
        // pickers show while it has none.
        // ###########################################################################################
        public sealed record Facts(string Name, int Rows, bool CanDelete, string Colors, int Going = 0)
        {
            public IReadOnlyList<string> Components { get; init; } = [];

            public int BoardFiles { get; init; }

            public int BoardLinks { get; init; }

            public string? OwnBackground { get; init; }

            public string? OwnText { get; init; }

            public Avalonia.Media.Color DefaultBackground { get; init; }

            public Avalonia.Media.Color DefaultText { get; init; }
        }

        private Facts thisFacts;

        // The colors as picked here, ahead of the draft until picking pauses - "#RRGGBB", null for none.
        private string? thisBackground;
        private string? thisText;
        private bool thisShowAll;

        public ComponentPaneRegionRow(Facts facts)
        {
            this.thisFacts = facts;
            this.thisBackground = facts.OwnBackground;
            this.thisText = facts.OwnText;
        }

        public string Name => this.thisFacts.Name;

        // ---- its colors (owner request, 2026-11-11) ----------------------------------------------------

        // Whether a property is one of the colors a person picks - what goes into the draft once picking pauses.
        public static bool IsColorChange(string? property) =>
            property is nameof(ComponentPaneRegionRow.BackgroundColor) or nameof(ComponentPaneRegionRow.TextColor);

        // The colors the board is to give it - null for the ones CRT gives it, and for a text worked out from the background.
        public string? OwnBackground => this.thisBackground;

        public string? OwnText => this.thisBackground is null ? null : this.thisText;

        public bool HasOwnColors => this.thisBackground is not null;

        public bool HasCrtColors => !this.HasOwnColors;

        public bool HasOwnText => this.OwnText is not null;

        // The background: the board's, else CRT's. Picking writes it - unless it is the color shown already, which a picker
        // writes back as it opens (CLAUDE.md: "A picker writing back the colour it shows writes nothing").
        public Avalonia.Media.Color BackgroundColor
        {
            get => ComponentPaneRegionRow.Parse(this.thisBackground) ?? this.thisFacts.DefaultBackground;
            set
            {
                if (value == this.BackgroundColor)
                {
                    return;
                }

                this.thisBackground = ComponentPaneRegionRow.Hex(value);
                this.RaiseColors();
            }
        }

        // The text: the one set, else - with a background of the board's own - the one that reads best on it
        // (RegionColors.ContrastText), else CRT's. Picked, it is set; picked while the region has CRT's colors, the background
        // shown becomes the board's too, as the pair is the region's colors.
        public Avalonia.Media.Color TextColor
        {
            get => ComponentPaneRegionRow.Parse(this.OwnText)
                   ?? (this.thisBackground is { } background ? ComponentPaneRegionRow.Parse(RegionColors.ContrastText(background)) : null)
                   ?? this.thisFacts.DefaultText;
            set
            {
                if (value == this.TextColor)
                {
                    return;
                }

                this.thisBackground ??= ComponentPaneRegionRow.Hex(this.thisFacts.DefaultBackground);
                this.thisText = ComponentPaneRegionRow.Hex(value);
                this.RaiseColors();
            }
        }

        public Avalonia.Media.IBrush BackgroundBrush => new Avalonia.Media.SolidColorBrush(this.BackgroundColor);

        public Avalonia.Media.IBrush TextBrush => new Avalonia.Media.SolidColorBrush(this.TextColor);

        // "Automatic text color": the text worked out from the background again.
        public void UseAutomaticText()
        {
            this.thisText = null;
            this.RaiseColors();
        }

        // "Reset colors": the colors CRT gives the region.
        public void UseCrtColors()
        {
            this.thisBackground = null;
            this.thisText = null;
            this.RaiseColors();
        }

        private void RaiseColors()
        {
            foreach (string property in new[]
                     {
                         nameof(this.BackgroundColor), nameof(this.TextColor), nameof(this.BackgroundBrush), nameof(this.TextBrush),
                         nameof(this.HasOwnColors), nameof(this.HasCrtColors), nameof(this.HasOwnText), nameof(this.OwnBackground),
                         nameof(this.OwnText),
                     })
            {
                this.Raise(property);
            }
        }

        private static Avalonia.Media.Color? Parse(string? hex) =>
            RegionColors.Normalised(hex) is { } color && Avalonia.Media.Color.TryParse(color, out Avalonia.Media.Color parsed) ? parsed : null;

        private static string Hex(Avalonia.Media.Color color) =>
            string.Create(System.Globalization.CultureInfo.InvariantCulture, $"#{color.R:X2}{color.G:X2}{color.B:X2}");

        // ---- what it is applied on (owner request, 2026-11-11) -------------------------------------------

        // The components it is applied on, as buttons opening each - the first ComponentsNamedAtMost unless all are asked for.
        public IReadOnlyList<ComponentPaneRegionComponent> ComponentsShown =>
            this.thisFacts.Components
                .Take(this.thisShowAll ? int.MaxValue : ComponentPaneRegionRow.ComponentsNamedAtMost)
                .Select(label => new ComponentPaneRegionComponent(label, this.Name))
                .ToList();

        public bool HasComponents => this.thisFacts.Components.Count > 0;

        // What the line under "Applied on:" holds: the components shown, then - while there are more than are named at first -
        // "... and <b>N</b> more" with "Show all" (or "Show fewer") as its last item, on the same line (owner request,
        // 2026-11-11).
        public IReadOnlyList<object> ChipsShown =>
            this.CanShowAll
                ? [.. this.ComponentsShown, new ComponentPaneRegionMore(this.MoreText, this.ShowAllText, this)]
                : [.. this.ComponentsShown];

        // "... and <b>4</b> more" while some are not named - nothing once all are.
        public string MoreText => this.HiddenComponents > 0
            ? $"... and <b>{this.HiddenComponents.ToString(System.Globalization.CultureInfo.InvariantCulture)}</b> more"
            : string.Empty;

        public bool HasMore => this.HiddenComponents > 0;

        // Offered whenever there are more than are named at first: "Show all" or, once shown, "Show fewer".
        public bool CanShowAll => this.thisFacts.Components.Count > ComponentPaneRegionRow.ComponentsNamedAtMost;

        public string ShowAllText => this.thisShowAll ? "Show fewer" : "Show all";

        private int HiddenComponents => this.thisShowAll ? 0 : Math.Max(0, this.thisFacts.Components.Count - ComponentPaneRegionRow.ComponentsNamedAtMost);

        public bool ShowAll
        {
            get => this.thisShowAll;
            set
            {
                if (this.thisShowAll == value)
                {
                    return;
                }

                this.thisShowAll = value;
                this.RaiseComponents();
            }
        }

        // "Also on <b>1</b> board file and <b>2</b> board links." - nothing when none.
        public string BoardRowsText => ComponentPaneWording.RegionOnBoardRows(this.thisFacts.BoardFiles, this.thisFacts.BoardLinks);

        public bool HasBoardRows => this.BoardRowsText.Length > 0;

        private void RaiseComponents()
        {
            foreach (string property in new[]
                     {
                         nameof(this.ComponentsShown), nameof(this.ChipsShown), nameof(this.HasComponents), nameof(this.MoreText), nameof(this.HasMore),
                         nameof(this.CanShowAll), nameof(this.ShowAllText), nameof(this.BoardRowsText), nameof(this.HasBoardRows),
                     })
            {
                this.Raise(property);
            }
        }

        public bool CanDelete => this.thisFacts.CanDelete;

        // The colors resource it is drawn in - "Schematics_Region_PAL", "Schematics_Region_Own2".
        public string Colors => this.thisFacts.Colors;

        public bool IsPal => this.Colors == "Schematics_Region_PAL";

        public bool IsNtsc => this.Colors == "Schematics_Region_NTSC";

        public bool IsOwn1 => this.Colors == "Schematics_Region_Own1";

        public bool IsOwn2 => this.Colors == "Schematics_Region_Own2";

        public bool IsOwn3 => this.Colors == "Schematics_Region_Own3";

        public bool IsOwn4 => this.Colors == "Schematics_Region_Own4";

        // Why it cannot be deleted, its count marked bold (BoldRuns) - "" when it can. How many rows say it is said here
        // alone (owner request, 2026-11-09: a count beside the name too was "double-wording").
        public string Why => this.thisFacts.CanDelete
            ? this.thisFacts.Going > 0 ? ComponentPaneWording.RegionTakesRows(this.thisFacts.Going) : string.Empty
            : this.thisFacts.Rows > 0 ? ComponentPaneWording.RegionInUse(this.thisFacts.Rows)
            : ComponentPaneWording.RegionsAtFewest;

        // Whether there is anything to say under it - why it stays, or what goes with it.
        public bool HasWhy => this.Why.Length > 0;

        public bool Shows(Facts facts) => string.Equals(this.thisFacts.Name, facts.Name, StringComparison.Ordinal);

        // The row's facts as they now are - its colors, whether it can go and why not.
        public void Use(Facts facts)
        {
            // The same facts: nothing raised, so its chips are not made again. Its components are compared by what they say -
            // HoldingOf hands a new list every time, and the record compared it by reference (code review, 2026-11-11).
            if (this.thisFacts == facts with { Components = this.thisFacts.Components } &&
                this.thisFacts.Components.SequenceEqual(facts.Components, StringComparer.Ordinal))
            {
                return;
            }

            this.thisFacts = facts;
            this.thisBackground = facts.OwnBackground;
            this.thisText = facts.OwnText;

            foreach (string property in new[]
                     {
                         nameof(this.CanDelete), nameof(this.Colors), nameof(this.IsPal), nameof(this.IsNtsc), nameof(this.IsOwn1),
                         nameof(this.IsOwn2), nameof(this.IsOwn3), nameof(this.IsOwn4), nameof(this.Why), nameof(this.HasWhy),
                     })
            {
                this.Raise(property);
            }

            this.RaiseColors();
            this.RaiseComponents();
        }
    }

    // One component a region is applied on, as a button on "Board: Regions" - opening it in that region.
    public sealed record ComponentPaneRegionComponent(string Label, string Region);

    // The end of that line: "... and <b>N</b> more" (empty once all are shown) and "Show all" or "Show fewer" for `Row`.
    public sealed record ComponentPaneRegionMore(string Text, string ShowAllText, ComponentPaneRegionRow Row)
    {
        public bool HasText => this.Text.Length > 0;
    }
}
