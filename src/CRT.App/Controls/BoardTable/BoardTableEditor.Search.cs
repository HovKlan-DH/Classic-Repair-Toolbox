using Avalonia.Controls;
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
    // THE MARKS are drawn by the table's own cell text, BoardTableCellText: the editor hands the
    // search, in the Workbooks tab's search colours (UseSearchColours), to the grid once, and every
    // cell's text inherits it and marks what it finds in its own row's text. ProDataGrid's own
    // search marks are not used - they showed another row's text in a recycled cell (owner report,
    // 2026-10-09; see BoardTableCellText). The cell keeps its own colour under a match.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        private static readonly TimeSpan SearchDelay = TimeSpan.FromMilliseconds(200);

        // The search applied, and the text it was applied from (BoardTableSearch.None for blank).
        private BoardTableSearch thisSearch = BoardTableSearch.None;
        private string thisSearchText = string.Empty;

        private DispatcherTimer? thisSearchTimer;

        // The Workbooks tab's search colours, read again on a theme change (UseSearchColours).
        private IBrush thisSearchHitBackground = Brushes.Khaki;
        private IBrush thisSearchHitForeground = Brushes.Black;

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
            this.SearchBox.TextChanged += this.OnSearchBoxTextChanged;
            this.ClearSearchButton.Click += this.OnClearSearchClick;

            this.UseSearchColours();
            this.ActualThemeVariantChanged += (_, _) => this.UseSearchColours();
        }

        // ###########################################################################################
        // *** THE MARKS ARE IN THE WORKBOOKS TAB'S OWN SEARCH COLOURS *** (Workbooks_SearchHit_Bg and
        // _Fg), read through ThemeResources - a control's own lookup does not reach a theme-variant
        // key - and again on a theme change, which hands the cells the search afresh.
        // ###########################################################################################
        private void UseSearchColours()
        {
            this.thisSearchHitBackground = ThemeResources.Resolve<IBrush>("Workbooks_SearchHit_Bg", Brushes.Khaki);
            this.thisSearchHitForeground = ThemeResources.Resolve<IBrush>("Workbooks_SearchHit_Fg", Brushes.Black);
            this.UpdateSearchMarks();
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
        // The search handed to every cell's text (BoardTableCellText.Marks, inherited from the grid),
        // which marks what it finds in its own row - a new one each time, so every cell looks again.
        // None while nothing is searched for.
        // ###########################################################################################
        private void UpdateSearchMarks()
        {
            BoardTableCellText.SetMarks(
                this.TableGrid,
                this.thisSearch.IsActive
                    ? new BoardTableSearchMarks(this.thisSearch, this.thisSearchHitBackground, this.thisSearchHitForeground)
                    : null);
        }
    }
}
