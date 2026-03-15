using Avalonia.Controls;
using Avalonia.Interactivity;
using DataHandling;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace CRT
{
    public partial class TabOverview : UserControl
    {
        private Main? _mainWindow;
        private List<OverviewRow> _allRows = new();

        public TabOverview()
        {
            this.InitializeComponent();
        }

        // ###########################################################################################
        // Initializes the overview tab with a reference to the main window.
        // ###########################################################################################
        public void Initialize(Main mainWindow)
        {
            this._mainWindow = mainWindow;
        }

        // ###########################################################################################
        // Populates the overview list based on the selected board data.
        // ###########################################################################################
        public void LoadData(BoardData boardData)
        {
            var rows = new List<OverviewRow>();

            foreach (var comp in boardData.Components)
            {
                var note = boardData.ComponentImages
                    .FirstOrDefault(ci => string.Equals(ci.BoardLabel, comp.BoardLabel, StringComparison.OrdinalIgnoreCase))?.Note ?? string.Empty;

                var links = new List<OverviewLink>();

                links.AddRange(boardData.ComponentLocalFiles
                    .Where(lf => string.Equals(lf.BoardLabel, comp.BoardLabel, StringComparison.OrdinalIgnoreCase))
                    .Select(lf => new OverviewLink(lf.Name, lf.File, OverviewLinkType.LocalFile)));

                links.AddRange(boardData.ComponentLinks
                    .Where(l => string.Equals(l.BoardLabel, comp.BoardLabel, StringComparison.OrdinalIgnoreCase))
                    .Select(l => new OverviewLink(l.Name, l.Url, OverviewLinkType.WebLink)));

                rows.Add(new OverviewRow
                {
                    Component = comp.BoardLabel ?? string.Empty,
                    TechnicalName = comp.TechnicalNameOrValue ?? string.Empty,
                    FriendlyName = comp.FriendlyName ?? string.Empty,
                    PartNumber = comp.PartNumber ?? string.Empty,
                    ShortDescription = comp.Description ?? string.Empty,
                    Notes = note,
                    Links = links
                });
            }

            this._allRows = rows;
            this.OverviewItemsControl.ItemsSource = this._allRows;
        }

        // ###########################################################################################
        // Filters the overview list based on the provided search term locally.
        // ###########################################################################################
        public void ApplyFilter(string searchTerm)
        {
            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                this.OverviewItemsControl.ItemsSource = this._allRows;
                return;
            }

            var terms = searchTerm.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var filtered = new List<OverviewRow>();

            foreach (var row in this._allRows)
            {
                var parts = new List<string>(3);
                if (!string.IsNullOrWhiteSpace(row.Component))
                    parts.Add(row.Component.Trim());
                if (!string.IsNullOrWhiteSpace(row.FriendlyName))
                    parts.Add(row.FriendlyName.Trim());
                if (!string.IsNullOrWhiteSpace(row.TechnicalName))
                    parts.Add(row.TechnicalName.Trim());

                if (parts.Count == 0) continue;

                string displayString = string.Join(" | ", parts);
                bool matches = true;

                foreach (var term in terms)
                {
                    if (displayString.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        matches = false;
                        break;
                    }
                }

                if (matches)
                {
                    filtered.Add(row);
                }
            }

            this.OverviewItemsControl.ItemsSource = filtered;
        }

        // ###########################################################################################
        // Opens the component info popup when clicking a component link.
        // ###########################################################################################
        private void OnComponentClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is OverviewRow row && this._mainWindow != null)
            {
                var parts = new List<string>(3);
                if (!string.IsNullOrWhiteSpace(row.Component))
                    parts.Add(row.Component.Trim());
                if (!string.IsNullOrWhiteSpace(row.FriendlyName))
                    parts.Add(row.FriendlyName.Trim());
                if (!string.IsNullOrWhiteSpace(row.TechnicalName))
                    parts.Add(row.TechnicalName.Trim());

                string displayText = string.Join(" | ", parts);
                this._mainWindow.OpenComponentInfoPopup(row.Component, displayText);
            }
        }

        // ###########################################################################################
        // Opens a link based on whether it is a local file or web URL.
        // ###########################################################################################
        // ###########################################################################################
        // Opens a link based on whether it is a local file or web URL.
        // ###########################################################################################
        private void OnLinkClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is OverviewLink link)
            {
                if (link.IsLocalFile)
                {
                    var fullPath = Path.Combine(DataManager.DataRoot, link.Target.Replace('/', Path.DirectorySeparatorChar));
                    try
                    {
                        Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Failed to open local file - [{fullPath}] - [{ex.Message}]");
                    }
                }
                else
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo(link.Target) { UseShellExecute = true });
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Failed to open web link - [{link.Target}] - [{ex.Message}]");
                    }
                }
            }
        }
    }

    public class OverviewRow
    {
        public string Component { get; init; } = string.Empty;
        public string TechnicalName { get; init; } = string.Empty;
        public string FriendlyName { get; init; } = string.Empty;
        public string PartNumber { get; init; } = string.Empty;
        public string ShortDescription { get; init; } = string.Empty;
        public string Notes { get; init; } = string.Empty;
        public List<OverviewLink> Links { get; init; } = new();
    }

    public enum OverviewLinkType
    {
        LocalFile,
        WebLink
    }

    public class OverviewLink
    {
        public string Name { get; }
        public string Target { get; }
        public OverviewLinkType Type { get; }

        public bool IsLocalFile => this.Type == OverviewLinkType.LocalFile;
        public bool IsWebLink => this.Type == OverviewLinkType.WebLink;

        public OverviewLink(string name, string target, OverviewLinkType type)
        {
            this.Name = name;
            this.Target = target;
            this.Type = type;
        }
    }

}