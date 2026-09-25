using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace CRT
{
    public partial class TabOverview : UserControl
    {
        private Main? _mainWindow;
        private List<OverviewRow> _allRows = new();
        private List<OverviewRow> _visibleRows = new();

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
                    IsSelectedForPrint = true,
                    Component = comp.BoardLabel ?? string.Empty,
                    Category = comp.Category ?? string.Empty,
                    TechnicalName = comp.TechnicalNameOrValue ?? string.Empty,
                    FriendlyName = comp.FriendlyName ?? string.Empty,
                    PartNumber = comp.PartNumber ?? string.Empty,
                    ShortDescription = comp.Description ?? string.Empty,
                    Notes = note,
                    Links = links
                });
            }

            this._allRows = rows;
            this._visibleRows = rows;
            this.OverviewItemsControl.ItemsSource = this._visibleRows;
        }

        // ###########################################################################################
        // Filters the overview list based on the active component/category filters and search term.
        // ###########################################################################################
        public void ApplyFilter(string searchTerm)
        {
            this._visibleRows = this.GetFilteredRows(searchTerm);
            this.OverviewItemsControl.ItemsSource = this._visibleRows;
        }

        // ###########################################################################################
        // Returns the overview rows that survive the current component/category/search filters.
        // ###########################################################################################
        private List<OverviewRow> GetFilteredRows(string searchTerm)
        {
            IEnumerable<OverviewRow> rows = this._allRows;

            var visibleBoardLabels = this.GetVisibleOverviewBoardLabels();
            var selectedBoardLabels = this.GetSelectedOverviewBoardLabels();

            if (selectedBoardLabels.Count > 0)
            {
                rows = rows.Where(row => selectedBoardLabels.Contains(row.Component));
            }
            else if (visibleBoardLabels.Count > 0)
            {
                rows = rows.Where(row => visibleBoardLabels.Contains(row.Component));
            }
            else if (this._mainWindow != null)
            {
                rows = Enumerable.Empty<OverviewRow>();
            }

            var terms = string.IsNullOrWhiteSpace(searchTerm)
                ? Array.Empty<string>()
                : searchTerm.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (terms.Length == 0)
            {
                return rows.ToList();
            }

            return rows
                .Where(row =>
                {
                    string displayString = OverviewHtmlBuilder.BuildOverviewDisplayString(row);
                    if (string.IsNullOrWhiteSpace(displayString))
                        return false;

                    foreach (var term in terms)
                    {
                        if (displayString.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                            return false;
                    }

                    return true;
                })
                .ToList();
        }

        // ###########################################################################################
        // Returns the currently visible component board labels from the main component list.
        // ###########################################################################################
        private HashSet<string> GetVisibleOverviewBoardLabels()
        {
            if (this._mainWindow == null)
            {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            return new HashSet<string>(
                this._mainWindow.ComponentFilterListBox.ItemsSource?
                    .Cast<ComponentListItem>()
                    .Where(item => !string.IsNullOrWhiteSpace(item.BoardLabel))
                    .Select(item => item.BoardLabel.Trim())
                ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Returns the currently selected component board labels from the main component list.
        // ###########################################################################################
        private HashSet<string> GetSelectedOverviewBoardLabels()
        {
            if (this._mainWindow == null)
            {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            return new HashSet<string>(
                this._mainWindow.ComponentFilterListBox.SelectedItems?
                    .Cast<ComponentListItem>()
                    .Where(item => !string.IsNullOrWhiteSpace(item.BoardLabel))
                    .Select(item => item.BoardLabel.Trim())
                ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Returns the rows currently visible in the overview and checked for printing.
        // ###########################################################################################
        private List<OverviewRow> GetPrintableRows()
        {
            IEnumerable<OverviewRow> rows = this._visibleRows.Count > 0
                ? this._visibleRows
                : (this.OverviewItemsControl.ItemsSource as IEnumerable<OverviewRow>) ?? Enumerable.Empty<OverviewRow>();

            return rows.Where(row => row.IsSelectedForPrint).ToList();
        }

        // ###########################################################################################
        // Selects or deselects every currently visible row for printing, driven by the header checkbox.
        // ###########################################################################################
        private void OnSelectAllPrintClick(object? sender, RoutedEventArgs e)
        {
            if (sender is not CheckBox checkBox)
            {
                return;
            }

            bool select = checkBox.IsChecked ?? false;

            foreach (var row in this._visibleRows)
            {
                row.IsSelectedForPrint = select;
            }
        }

        // ###########################################################################################
        // Opens a printable HTML document for the currently visible overview rows.
        // ###########################################################################################
        private void OnPrintComponentListClick(object? sender, RoutedEventArgs e)
        {
            var printableRows = this.GetPrintableRows();
            if (printableRows.Count == 0)
            {
                Logger.Info("Overview print skipped because there are no rows checked for printing");
                return;
            }

            try
            {
                string tempFilePath = Path.Combine(
                    Path.GetTempPath(),
                    $"crt-overview-print-{Guid.NewGuid():N}.html");

                File.WriteAllText(tempFilePath, OverviewHtmlBuilder.BuildPrintableHtml(printableRows), Encoding.UTF8);

                Process.Start(new ProcessStartInfo(tempFilePath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to print overview component list - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Opens a printable HTML document with the bill of materials for the currently visible rows.
        // ###########################################################################################
        private void OnPrintComponentQuantitiesClick(object? sender, RoutedEventArgs e)
        {
            var printableRows = this.GetPrintableRows();
            if (printableRows.Count == 0)
            {
                Logger.Info("Overview print skipped because there are no rows checked for printing");
                return;
            }

            var groups = OverviewHtmlBuilder.BuildQuantityGroups(printableRows);

            try
            {
                string tempFilePath = Path.Combine(
                    Path.GetTempPath(),
                    $"crt-overview-print-bom-{Guid.NewGuid():N}.html");

                File.WriteAllText(tempFilePath, OverviewHtmlBuilder.BuildPrintableQuantitiesHtml(groups), Encoding.UTF8);

                Process.Start(new ProcessStartInfo(tempFilePath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to print overview bill of materials - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Subscribes each row to pointer presses at the tunnel phase, so the click is seen before any
        // child control (button, checkbox, selectable text) has a chance to mark it as handled.
        // ###########################################################################################
        private void OnOverviewRowLoaded(object? sender, RoutedEventArgs e)
        {
            if (sender is Border border)
            {
                border.AddHandler(InputElement.PointerPressedEvent, this.OnOverviewRowPointerPressed, RoutingStrategies.Tunnel);
            }
        }

        // ###########################################################################################
        // Toggles the row's print-selection checkbox when clicking anywhere in the row other than a
        // button or the checkbox itself (e.g. the component link or a file/web link).
        // ###########################################################################################
        private void OnOverviewRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Border border || border.DataContext is not OverviewRow row)
                return;

            if (e.Source is Visual clickedVisual &&
                (clickedVisual.FindAncestorOfType<Button>(includeSelf: true) != null ||
                 clickedVisual.FindAncestorOfType<CheckBox>(includeSelf: true) != null))
            {
                return;
            }

            row.IsSelectedForPrint = !row.IsSelectedForPrint;
        }

        // ###########################################################################################
        // Opens the component info popup when clicking a component link.
        // ###########################################################################################
        private void OnComponentClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is OverviewRow row && this._mainWindow != null)
            {
                string displayText = OverviewHtmlBuilder.BuildOverviewDisplayString(row);
                this._mainWindow.OpenComponentInfoPopup(row.Component, displayText);
            }
        }

        // ###########################################################################################
        // Opens a link based on whether it is a local file or web URL.
        //
        // *** BOTH GO THROUGH ExternalTargetLauncher, never a raw Process.Start. *** This handed the
        // board's own strings straight to the OS shell, which runs a ".bat" or follows a ".lnk"
        // rather than displaying it - and once local files began resolving against the Drafts tree,
        // which a contributor edits by hand by design, a row naming "tool.bat" in the draft folder
        // would have been executed on click. The component popup's identical button already went
        // through the launcher and refused it; the two now behave the same.
        //
        // The launcher is handed the ROOT the file was found under (DraftFileResolver.ResolveWithSource),
        // because its containment check refuses a path outside the root it is given and a drafted
        // file is never under the data root.
        // ###########################################################################################
        private void OnLinkClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is OverviewLink link)
            {
                (string target, string? root) = this.ResolveLinkTarget(link);

                ExternalTargetLauncher.TryOpen(target, root);
            }
        }

        // ###########################################################################################
        // What a link click opens, and the containment root to open it against. A web link is
        // passed through unchanged with no root - the launcher decides whether its scheme is one it
        // will open. A local file resolves draft-first through the VIEWED draft folder, so "view
        // boards as officially published" opens the published copy.
        // ###########################################################################################
        internal (string Target, string? Root) ResolveLinkTarget(OverviewLink link)
        {
            if (!link.IsLocalFile)
            {
                return (link.Target, null);
            }

            string draftSystemFolder = this._mainWindow?.GetCurrentBoardEntry() is { } entry
                ? DraftManager.GetViewedSystemFolder(entry.ExcelDataFile)
                : string.Empty;

            return TabOverview.ResolveLocalFileTarget(DataManager.DataRoot, draftSystemFolder, link.Target);
        }

        // Pure apart from File.Exists, so the root pairing is testable without a window.
        internal static (string Target, string? Root) ResolveLocalFileTarget(
            string dataRoot,
            string draftSystemFolder,
            string storedFile)
        {
            DraftFileResolution? resolved = DraftFileResolver.ResolveWithSource(dataRoot, draftSystemFolder, storedFile);

            return resolved is not null
                ? (resolved.FullPath, resolved.Root)
                : (Path.Combine(dataRoot, storedFile.Replace('/', Path.DirectorySeparatorChar)), dataRoot);
        }
    }

}