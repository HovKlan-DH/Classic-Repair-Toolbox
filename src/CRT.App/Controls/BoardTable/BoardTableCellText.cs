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
            IReadOnlyList<(string Text, bool IsMatch)> segments = marks is null ? [] : marks.Search.Query.SplitIntoSegments(text);

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
                runs.Add(isMatch
                    ? new Run(part) { Background = marks!.Background, Foreground = marks.Foreground }
                    : new Run(part));
            }

            this.Text = null;
            this.Inlines = runs;
        }
    }

    // The search the table applies, and the colours its finds are drawn in. A new one for every
    // search (and every theme change), so the cells look again.
    internal sealed record BoardTableSearchMarks(BoardTableSearch Search, IBrush Background, IBrush Foreground);

    // ###########################################################################################
    // A data column of the table: the grid's own text column - its editing, its cell theme - with
    // BoardTableCellText as what a cell shows when it is not being edited. `DisplayPath` is the
    // cell's text, read one way for showing it (the column's Binding stays two-way, for editing).
    // ###########################################################################################
    internal sealed class BoardTableTextColumn : DataGridTextColumn
    {
        public required string DisplayPath { get; init; }

        protected override Control GenerateElement(DataGridCell cell, object dataItem)
        {
            // The grid's own drawn or direct cells are not used here; if one ever is, it keeps its
            // own content.
            if (cell.GetType() != typeof(DataGridCell))
            {
                return base.GenerateElement(cell, dataItem);
            }

            var text = new BoardTableCellText { Name = "CellTextBlock" };

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
