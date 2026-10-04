using Avalonia.Controls;
using Avalonia.Controls.DataGridSearching;
using System;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // HOW THE TABLE'S SEARCH MARKS REACH THE CELLS (2026-10-02, BoardTableEditor.Search.cs).
    //
    // ProDataGrid draws every cell's text with its own search-aware text block, which marks the
    // runs its search model's RESULTS name. The grid's own search would decide those results with
    // its own grammar; the table's search is the Workbooks tab's (BoardTableSearch), so the results
    // are worked out by the editor and handed over.
    //
    // *** THIS ADAPTER IS WHAT KEEPS THEM. *** The grid's stock adapter recomputes the results
    // whenever the rows change - a filter, an insert, a refresh - and with its own grammar that
    // cleared every mark the editor had set (found by a spike, 2026-10-02: the marks went at the
    // first view refresh). This one answers every such recompute with the editor's results.
    // ###########################################################################################
    internal sealed class BoardTableSearchAdapterFactory(Func<IReadOnlyList<SearchResult>> resultsOnScreen) : IDataGridSearchAdapterFactory
    {
        public DataGridSearchAdapter Create(DataGrid grid, ISearchModel model) =>
            new BoardTableSearchAdapter(model, () => grid.Columns, resultsOnScreen);
    }

    internal sealed class BoardTableSearchAdapter(
        ISearchModel model,
        Func<IEnumerable<DataGridColumn>> columns,
        Func<IReadOnlyList<SearchResult>> resultsOnScreen) : DataGridSearchAdapter(model, columns)
    {
        protected override bool TryApplyModelToView(
            IReadOnlyList<SearchDescriptor> descriptors,
            IReadOnlyList<SearchDescriptor> previousDescriptors,
            out IReadOnlyList<SearchResult> results)
        {
            results = resultsOnScreen();
            return true;
        }
    }
}
