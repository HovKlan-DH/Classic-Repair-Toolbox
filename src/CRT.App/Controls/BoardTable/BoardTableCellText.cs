using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Data;
using Avalonia.Media;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE TEXT OF A TABLE CELL, WITH THE SEARCH'S FINDS MARKED (owner report, 2026-10-09: a search
    // for "rep" left "replica" unmarked in two rows - which turned out to be two rows showing ANOTHER
    // row's text).
    //
    // *** WHY NOT PRODATAGRID'S OWN SEARCH TEXT. *** Its DataGridSearchTextBlock builds the marked
    // runs from the search RESULT's text, kept on the block (SearchText) and preferred to the block's
    // own. The grid recycles a row's cells for another row as the view changes - a search narrows it,
    // a scroll moves it - and the new row's text arrives while the old result is still on the block,
    // so the old row's text was taken as the text; with no match in the new row the block then fell
    // back to that remembered text. "Schematics #1 of 2" showed as "Top (replica)", unmarked (pinned
    // by BoardTableEditorSearchTests).
    //
    // So the table draws its own: the row's text is bound to CellText, never to Text, and the runs
    // are worked out from CellText and the search the editor hands every cell (Marks, inherited from
    // the grid) each time either changes - nothing is remembered between rows. With no search, or
    // nothing found, it is one plain Text, which lays out faster than runs.
    // ###########################################################################################
    internal sealed class BoardTableCellText : TextBlock
    {
        public static readonly StyledProperty<string?> CellTextProperty =
            AvaloniaProperty.Register<BoardTableCellText, string?>(nameof(CellText));

        // The search and its colours, set once on the grid and inherited by every cell's text.
        public static readonly AttachedProperty<BoardTableSearchMarks?> MarksProperty =
            AvaloniaProperty.RegisterAttached<BoardTableCellText, Control, BoardTableSearchMarks?>("Marks", inherits: true);

        public string? CellText
        {
            get => this.GetValue(CellTextProperty);
            set => this.SetValue(CellTextProperty, value);
        }

        // The sheet's column it shows - a column the replace bar's "Columns" left out is never marked (BoardTableSearch
        // .Searches). -1 for a cell of no column of the sheet.
        public int ColumnIndex { get; init; } = -1;

        public static void SetMarks(Control control, BoardTableSearchMarks? marks) => control.SetValue(MarksProperty, marks);

        public static BoardTableSearchMarks? GetMarks(Control control) => control.GetValue(MarksProperty);

        protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
        {
            base.OnPropertyChanged(change);

            if (change.Property == CellTextProperty || change.Property == MarksProperty)
            {
                this.ShowText();
            }
        }

        // The runs - split by the search's own SplitIntoSegments (CRT.Data, tested), the split the
        // Workbooks tab marks its text with, so the two never disagree on where a find starts.
        private void ShowText()
        {
            string text = this.CellText ?? string.Empty;
            BoardTableSearchMarks? marks = this.GetValue(MarksProperty);

            // A column "Find" does not look in shows what it holds, unmarked - nothing there will be replaced.
            if (marks is not null && this.ColumnIndex >= 0 && this.DataContext is BoardTableRow row && !marks.Search.Searches(row.Sheet, this.ColumnIndex))
            {
                marks = null;
            }

            // The replace bar's "Find" is cut where "Replace all" replaces (BoardTableReplaceText) - never the search's
            // merged finds, which mark more than a replace touches.
            IReadOnlyList<(string Text, bool IsMatch)> segments =
                marks is null ? []
                : marks.IsLiteral ? BoardTableReplaceText.Split(text, marks.Search.Text)
                : marks.Search.Query.SplitIntoSegments(text);

            // Runs first cleared, always: a block that was marked keeps its runs otherwise, and they
            // would draw under the new text.
            this.Inlines?.Clear();

            if (!segments.Any(segment => segment.IsMatch))
            {
                this.Text = text;
                return;
            }

            var runs = new InlineCollection();

            foreach ((string part, bool isMatch) in segments)
            {
                // What "Replace with" holds, drawn in place of what it replaces (owner request, 2026-11-07: "so I can see
                // directly what the new text will look like").
                runs.Add(!isMatch ? new Run(part)
                    : marks!.Replacement is { } replacement
                        ? new Run(replacement) { Background = marks.ReplacementBackground, Foreground = marks.ReplacementForeground }
                        : new Run(part) { Background = marks.Background, Foreground = marks.Foreground });
            }

            this.Text = null;
            this.Inlines = runs;
        }
    }

    // ###########################################################################################
    // The search the table applies, and the colours its finds are drawn in. A new one for every
    // search (and every theme change), so the cells look again. The replace bar's "Find" is
    // `IsLiteral`, and with "Replace with" typed carries the `Replacement` drawn in its finds' place, in
    // its own colors.
    // ###########################################################################################
    internal sealed record BoardTableSearchMarks(BoardTableSearch Search, IBrush Background, IBrush Foreground)
    {
        public bool IsLiteral { get; init; }

        public string? Replacement { get; init; }

        public IBrush? ReplacementBackground { get; init; }

        public IBrush? ReplacementForeground { get; init; }
    }

    // ###########################################################################################
    // A data column of the table: the grid's own text column - its editing, its cell theme - with
    // BoardTableCellText as what a cell shows when it is not being edited. `DisplayPath` is the
    // cell's text, read one way for showing it (the column's Binding stays two-way, for editing).
    // ###########################################################################################
    internal sealed class BoardTableTextColumn : DataGridTextColumn
    {
        public required string DisplayPath { get; init; }

        // The sheet's column this is - the cells' text is told it (BoardTableCellText.ColumnIndex).
        public int ColumnIndex { get; init; } = -1;

        protected override Control GenerateElement(DataGridCell cell, object dataItem)
        {
            // The grid's own drawn or direct cells are not used here; if one ever is, it keeps its
            // own content.
            if (cell.GetType() != typeof(DataGridCell))
            {
                return base.GenerateElement(cell, dataItem);
            }

            var text = new BoardTableCellText { Name = "CellTextBlock", ColumnIndex = this.ColumnIndex };

            if (this.CellTextBlockTheme is { } theme)
            {
                text.Theme = theme;
            }

            if (this.IsSet(DataGridTextColumn.FontSizeProperty))
            {
                text.FontSize = this.FontSize;
            }

            text.Bind(BoardTableCellText.CellTextProperty, new Binding(this.DisplayPath) { Mode = BindingMode.OneWay });

            return text;
        }
    }
}
