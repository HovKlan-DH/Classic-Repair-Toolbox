using Avalonia.Controls;
using Avalonia.Controls.DataGridSearching;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE SEARCH BOX above the table (owner request, 2026-10-02: "how about the same search/filter
    // as we have in the "Workbooks" tab, where it highlights what I search for").
    //
    // Which rows show is BoardTableSearch's (CRT.Data) - the Workbooks tab's own grammar; this part
    // only wires the box, filters the view and paints the marks. Typing waits a moment before it
    // applies, as the Workbooks tab's box does, so a word is one refilter rather than one per key.
    //
    // *** WHEN IT IS EMPTIED (agreed case D). *** It stays across sheet tabs, across CRT's own tabs
    // and through a save or Reload (which read the same draft back); closing the table (Clear) and
    // opening another draft (Load of another file) empty it. In the Maintainer tab, choosing another
    // submission closes the table first, so it is emptied there too.
    //
    // *** A SAVE OR RELOAD APPLIES IT AGAIN. *** The rows read back are new objects, so the search
    // is worked out afresh on them - a row edited so it no longer matches, which stayed on screen
    // (agreed case C), is then hidden.
    //
    // THE MARKS are the grid's own: each cell's text is drawn by ProDataGrid's search-aware text
    // block, which marks the runs named in the grid's search model. The editor works the runs out
    // with BoardTableSearch and hands them over through BoardTableSearchAdapter, in the Workbooks
    // tab's search colours (UseSearchColours). The cell themes keep each cell's own colour
    // under a match (BuildCellTheme) - the grid's look would fill a matching cell with its accent.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        private static readonly TimeSpan SearchDelay = TimeSpan.FromMilliseconds(200);

        // The search applied, and the text it was applied from (BoardTableSearch.None for blank).
        private BoardTableSearch thisSearch = BoardTableSearch.None;
        private string thisSearchText = string.Empty;

        private DispatcherTimer? thisSearchTimer;

        // Whether the table shows only some rows - a pill picked or a search typed. Rows cannot be
        // moved then: a position among rows that cannot be seen means nothing.
        private bool IsNarrowed => this.thisFilter != BoardTableRowKinds.None || this.thisSearch.IsActive;

        // The search applied - set to apply one at once (the box shows it too).
        internal string SearchText
        {
            get => this.thisSearchText;
            set
            {
                value ??= string.Empty;

                if (!string.Equals(this.SearchBox.Text ?? string.Empty, value, StringComparison.Ordinal))
                {
                    this.SearchBox.Text = value;
                }

                this.ApplySearch(value);
            }
        }

        // What the box holds, applied now rather than after the pause - for tests, which may build
        // the editor without ever showing it (a detached TextBox raises no TextChanged).
        internal void ApplyPendingSearchForTests() => this.ApplySearch(this.SearchBox.Text);

        private void WireSearch()
        {
            this.TableGrid.SearchAdapterFactory = new BoardTableSearchAdapterFactory(this.SearchResultsOnScreen);

            this.SearchBox.TextChanged += this.OnSearchBoxTextChanged;
            this.ClearSearchButton.Click += this.OnClearSearchClick;

            this.UseSearchColours();
            this.ActualThemeVariantChanged += (_, _) => this.UseSearchColours();
        }

        // ###########################################################################################
        // *** THE MARKS ARE IN THE WORKBOOKS TAB'S OWN SEARCH COLOURS *** (Workbooks_SearchHit_Bg and
        // _Fg), under the keys the grid's cell text looks for. Copied into this control's resources
        // in code, and again on a theme change: the grid looks its brushes up with a control's own
        // lookup, which does not reach a theme-variant key (see ThemeResources) - written into the
        // theme dictionaries, the marks came out in the grid's default blue.
        //
        // *** AND A MATCHING ROW IS NOT SHADED. *** The grid also lays a tint of the mark's colour
        // over every row with a match (seen in a render, 2026-10-02) - which turned every white
        // cell on screen pale yellow, since a search shows only rows with a match, and read as one
        // more row colour beside green, orange and red. Its two opacities are set to 0.
        // ###########################################################################################
        private void UseSearchColours()
        {
            IBrush background = ThemeResources.Resolve<IBrush>("Workbooks_SearchHit_Bg", Brushes.Khaki);

            this.Resources["DataGridSearchMatchBrush"] = background;
            this.Resources["DataGridSearchCurrentBrush"] = background;
            this.Resources["DataGridSearchMatchForegroundBrush"] = ThemeResources.Resolve<IBrush>("Workbooks_SearchHit_Fg", Brushes.Black);
            this.Resources["DataGridRowSearchMatchOpacity"] = 0d;
            this.Resources["DataGridRowSearchCurrentOpacity"] = 0d;
        }

        private void OnSearchBoxTextChanged(object? sender, TextChangedEventArgs e)
        {
            string text = this.SearchBox.Text ?? string.Empty;
            this.ClearSearchButton.IsVisible = text.Length > 0;

            if (string.Equals(text, this.thisSearchText, StringComparison.Ordinal))
            {
                this.thisSearchTimer?.Stop();
                return;
            }

            if (this.thisSearchTimer is null)
            {
                this.thisSearchTimer = new DispatcherTimer { Interval = BoardTableEditor.SearchDelay };
                this.thisSearchTimer.Tick += (_, _) => this.ApplySearch(this.SearchBox.Text);
            }

            this.thisSearchTimer.Stop();
            this.thisSearchTimer.Start();
        }

        // The cross inside the box: an empty search at once, and the box keeps the focus.
        private void OnClearSearchClick(object? sender, RoutedEventArgs e)
        {
            e.Handled = true;
            this.SearchText = string.Empty;
            this.SearchBox.Focus();
        }

        // Applies `text` as the search: which rows show, which tabs show (moving to a sheet with a
        // match when the one on screen has none, as a picked pill does), and the marks.
        private void ApplySearch(string? text)
        {
            this.thisSearchTimer?.Stop();

            text ??= string.Empty;
            this.thisSearchText = text;
            this.ClearSearchButton.IsVisible = text.Length > 0;

            this.TableGrid.CommitEdit();
            this.thisSearch = this.thisDocument is null ? BoardTableSearch.None : BoardTableSearch.For(this.thisDocument, text);

            this.UpdateRowsDraggable();
            this.ApplyViewFilter();
            this.ShowASheetWithRowsShown();
            this.UpdateSheetTabs();
            this.UpdateToolbar();
            this.UpdateSearchMarks();
        }

        // The search worked out afresh on the document just attached - its rows are new objects.
        private void ReapplySearchTo(BoardTableDocument document) =>
            this.thisSearch = BoardTableSearch.For(document, this.thisSearchText);

        // ###########################################################################################
        // The runs to mark, handed to the grid's search model. The model names the search with a
        // descriptor (its text - a new text is what makes the adapter ask again) and is given the
        // results directly as well, since an edit changes the runs without changing the text.
        // ###########################################################################################
        private void UpdateSearchMarks()
        {
            if (this.TableGrid.SearchModel is not { } model)
            {
                return;
            }

            // The found TEXT is marked, not only its cell - the grid's default marks the cell alone,
            // whose fill the cell theme takes back. And a match is never "current": no cell is
            // singled out, and the grid never moves the cursor or the selection to one.
            model.HighlightMode = SearchHighlightMode.TextAndCell;
            model.HighlightCurrent = false;
            model.UpdateSelectionOnNavigate = false;

            if (!this.thisSearch.IsActive)
            {
                if (model.Descriptors.Count > 0)
                {
                    model.Clear();
                }

                if (model.Results.Count > 0)
                {
                    model.UpdateResults([]);
                }

                return;
            }

            model.SetOrUpdate(new SearchDescriptor(
                this.thisSearchText,
                SearchMatchMode.Contains,
                SearchTermCombineMode.All,
                SearchScope.AllColumns,
                null,
                StringComparison.OrdinalIgnoreCase,
                CultureInfo.InvariantCulture,
                false,
                false,
                false,
                false));

            model.UpdateResults(this.SearchResultsOnScreen());
        }

        // Every run the search found in the cells of the rows on screen - the marker column and
        // hidden rows have none.
        private List<SearchResult> SearchResultsOnScreen()
        {
            var results = new List<SearchResult>();

            if (!this.thisSearch.IsActive || this.thisView is null)
            {
                return results;
            }

            List<DataGridColumn> columns = this.TableGrid.Columns.Where(column => column.Tag is int index && index >= 0).ToList();
            int rowIndex = 0;

            foreach (object item in this.thisView)
            {
                if (item is BoardTableRow row)
                {
                    foreach (DataGridColumn column in columns)
                    {
                        int index = (int)column.Tag!;
                        if (index >= row.Cells.Count)
                        {
                            continue;
                        }

                        string text = row.Cells[index].Text;
                        IReadOnlyList<WorklogSearchHit> hits = this.thisSearch.HitsIn(text);

                        if (hits.Count > 0)
                        {
                            results.Add(new SearchResult(
                                row,
                                rowIndex,
                                column,
                                column.DisplayIndex,
                                text,
                                hits.Select(hit => new SearchMatch(hit.Start, hit.Length)).ToList()));
                        }
                    }
                }

                rowIndex++;
            }

            return results;
        }
    }
}
