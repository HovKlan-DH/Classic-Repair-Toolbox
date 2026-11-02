using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // "BOARD: REGIONS" AND "BOARD: CATEGORIES" (owner requests, 2026-11-01), two of the board's own parts.
    //
    //   - REGIONS: the one tick "This board comes in PAL and NTSC versions" (BoardRegions - stored in the
    //     board file and published). It cannot be taken away while a row says a region: such a row would go
    //     on being there with no button left to reach it.
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

        private void ShowRegionsPart(BoardData board)
        {
            int rows = BoardRegions.RegionalRows(board);

            this.RegionsCheckBox.Content = ComponentPaneWording.RegionsHasOnBoard;
            this.RegionsCheckBox.IsChecked = BoardRegions.Has(board);
            this.RegionsCheckBox.IsEnabled = rows == 0;
            this.RegionsInUseText.Text = rows > 0 ? ComponentPaneWording.RegionsInUse(rows) : string.Empty;
            this.RegionsInUseText.IsVisible = rows > 0;
        }

        private void OnRegionsCheckClick(object? sender, RoutedEventArgs e) => this.SetRegions(this.RegionsCheckBox.IsChecked == true);

        internal void SetRegions(bool has)
        {
            if (this.thisSession is { } session)
            {
                this.Apply(DraftStepWording.Regions, board => BoardRegions.WithRegions(board, has));
                this.ShowRegionsPart(session.Content.Board);
                this.ShowBoardPartChanged();
            }
        }

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

        // What a category box offers: the board's, and those added and waiting for their first component.
        private List<string> CategoryChoices(BoardData board) =>
            ComponentListBuilder.BuildDistinctCategories(board)
                .Concat(this.HeldCategories(board))
                .OrderBy(category => category, StringComparer.OrdinalIgnoreCase)
                .ToList();

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

        public string NameDraft { get => this.thisNameDraft; set => this.Set(ref this.thisNameDraft, value); }

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
}
