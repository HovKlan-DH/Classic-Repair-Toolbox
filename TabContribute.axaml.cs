using Avalonia.Controls;
using Avalonia.Interactivity;
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
            if (boardData == null)
            {
                this.CategoriesHeaderControl.ItemsSource = null;
                this.CategoriesItemsControl.ItemsSource = null;
                return;
            }

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

            this.CategoriesHeaderControl.ItemsSource = columns;
            this.CategoriesItemsControl.ItemsSource = columns;
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