using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    public partial class TabContribute : UserControl
    {
        private Main? thisMainWindow;

        public TabContribute()
        {
            this.InitializeComponent();
        }

        // ###########################################################################################
        // Initializes the control with a reference to the main window.
        // ###########################################################################################
        public void Initialize(Main mainWindow)
        {
            this.thisMainWindow = mainWindow;
        }

        // ###########################################################################################
        // Loads the board data and populates the category columns with clickable components.
        // ###########################################################################################
        public void LoadData(BoardData? boardData, string region)
        {
            // Nothing can be contributed - not even a new component - until a board is loaded.
            this.AddNewComponentButton.IsEnabled = boardData != null;
            this.EditBoardAsDraftButton.IsEnabled = boardData != null;

            // Why "Edit board as draft" made no draft is about the board it was pressed on - never
            // left standing under another one (code review, 2026-10-09).
            this.ShowEditAsDraftProblem(null);

            if (boardData == null)
            {
                this.UpdateRevisionDate(null);
                this.CategoriesItemsControl.ItemsSource = null;
                return;
            }

            this.UpdateRevisionDate(boardData.RevisionDate);

            var componentsList = boardData.Components
                .Where(c =>
                    string.IsNullOrWhiteSpace(c.Region) ||
                    string.Equals(c.Region.Trim(), region, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var distinctCategories = new List<string>();
            var groupedItems = new Dictionary<string, List<ContributeComponentItem>>(StringComparer.OrdinalIgnoreCase);
            var seenByCategory = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var component in componentsList)
            {
                var category = string.IsNullOrWhiteSpace(component.Category)
                    ? "Uncategorized"
                    : component.Category.Trim();

                if (!groupedItems.ContainsKey(category))
                {
                    groupedItems[category] = new List<ContributeComponentItem>();
                    seenByCategory[category] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    distinctCategories.Add(category);
                }

                var boardLabel = component.BoardLabel?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(boardLabel) || !seenByCategory[category].Add(boardLabel))
                {
                    continue;
                }

                var tooltipParts = new List<string>();
                if (!string.IsNullOrWhiteSpace(boardLabel))
                    tooltipParts.Add(boardLabel);
                if (!string.IsNullOrWhiteSpace(component.FriendlyName))
                    tooltipParts.Add(component.FriendlyName.Trim());
                if (!string.IsNullOrWhiteSpace(component.TechnicalNameOrValue))
                    tooltipParts.Add(component.TechnicalNameOrValue.Trim());

                groupedItems[category].Add(new ContributeComponentItem
                {
                    BoardLabel = boardLabel,
                    DisplayText = boardLabel,
                    ToolTipText = string.Join(" | ", tooltipParts)
                });
            }

            var columns = distinctCategories
                .Select(cat => new CategoryColumn
                {
                    CategoryName = cat,
                    Components = groupedItems[cat]
                })
                .ToList();

            this.CategoriesItemsControl.ItemsSource = columns;
        }

        // ###########################################################################################
        // Shows or hides the board revision date text for the currently loaded board.
        // ###########################################################################################
        private void UpdateRevisionDate(string? revisionDate)
        {
            if (string.IsNullOrWhiteSpace(revisionDate))
            {
                this.RevisionDateText.Text = string.Empty;
                this.RevisionDatePanel.IsVisible = false;
                return;
            }

            this.RevisionDateText.Text = revisionDate;
            this.RevisionDatePanel.IsVisible = true;
        }

        // ###########################################################################################
        // Opens the maximized contribution editor for the clicked component.
        // ###########################################################################################
        private void OnComponentClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributeComponentItem item })
            {
                this.thisMainWindow?.OpenComponentContributionWindow(item.BoardLabel);
            }
        }

        // ###########################################################################################
        // Opens the same contribution editor on a component that is not in the board data at all,
        // so a contributor can suggest one the board is missing rather than only correcting the
        // components already listed above.
        // ###########################################################################################
        private void OnAddNewComponentClick(object? sender, RoutedEventArgs e)
        {
            this.thisMainWindow?.OpenNewComponentContributionWindow();
        }

        // ###########################################################################################
        // "Edit board as draft" - the board shown as the contributor's own draft, in the Drafts tab's
        // table; the draft there is when it has one (Main.EditBoardAsDraft.cs). Fire and forget, as
        // "Add a new board" below.
        // ###########################################################################################
        private void OnEditBoardAsDraftClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisMainWindow == null)
            {
                return;
            }

            _ = this.thisMainWindow.EditBoardAsDraftAsync();
        }

        // Why "Edit board as draft" made no draft - or nothing, to clear it.
        internal void ShowEditAsDraftProblem(string? problem)
        {
            this.EditAsDraftProblemText.Text = problem ?? string.Empty;
            this.EditAsDraftProblemText.IsVisible = !string.IsNullOrWhiteSpace(problem);
        }

        // ###########################################################################################
        // Creates a whole new hardware/board of the user's own as a local draft (session 2c, task 9)
        // - the board that is not in the list at all, rather than the component that is not on the
        // board. Unlike "Add new component" this needs no loaded board, which is the point.
        // ###########################################################################################
        private void OnAddNewBoardClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisMainWindow == null)
            {
                return;
            }

            // Fire and forget: the dialog owns the rest of the sequence, and awaiting here would
            // mean an async void handler for no benefit - the same shape the worklog buttons use.
            _ = this.thisMainWindow.OpenNewBoardWindowAsync();
        }
    }

    public class ContributeComponentItem
    {
        public string BoardLabel { get; init; } = string.Empty;
        public string DisplayText { get; init; } = string.Empty;
        public string ToolTipText { get; init; } = string.Empty;
    }

    public class CategoryColumn
    {
        public string CategoryName { get; init; } = string.Empty;
        public List<ContributeComponentItem> Components { get; init; } = new();
    }
}