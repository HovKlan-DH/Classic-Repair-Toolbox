using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace CRT
{
    // ###########################################################################################
    // HOVERING A CHANGED COLOR CELL SHOWS BOTH COLORS (owner request, 2026-11-01: "In 'Maintainer' tab: If
    // the color has changed, then I would like a popup showing the colors 'Changed from' and 'Changed
    // to'"). Text alone - "DeepSkyBlue" against "#1E90FF" - does not say how different two colors are.
    //
    // The file hover card's own popup, so it opens, moves and closes as that does
    // (BoardTableEditor.FilePreview.cs); the same in the Contribute tab and in the Maintainer tab, which is
    // why it is here and not in either. Which cells, and their two sides, are BoardTableColourCells'.
    // Part of the BoardTableEditor partial class - see BoardTableEditor.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        internal const string ColourChangedFrom = "Changed from";
        internal const string ColourChangedTo = "Changed to";

        // What a side with no color of its own says - CRT draws it in its default.
        internal const string ColourNotSet = "(not set)";

        // Whether this column's cells show the card rather than their text tooltip, which says the same.
        private bool ShowsColourChangesIn(int columnIndex) =>
            this.thisCurrentSheet is not null &&
            columnIndex >= 0 &&
            columnIndex < this.thisCurrentSheet.Columns.Count &&
            BoardTableColourCells.IsColourColumn(this.thisCurrentSheet.Name, this.thisCurrentSheet.Columns[columnIndex]);

        // The row and data-column index of a grid cell - only when it is a color cell that changed.
        private static (BoardTableRow Row, int Column)? ColourChangeCellOf(DataGridCell cell)
        {
            if (DataGridColumn.GetColumnContainingElement(cell)?.Tag is not int column ||
                cell.DataContext is not BoardTableRow row ||
                column < 0 ||
                column >= row.Cells.Count ||
                BoardTableColourCells.Of(row.Cells[column]) is null)
            {
                return null;
            }

            return (row, column);
        }

        // The card for one cell - null for any but a color cell that changed.
        internal static Control? BuildColourChangeCard(BoardTableRow row, int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= row.Cells.Count || BoardTableColourCells.Of(row.Cells[columnIndex]) is not { } change)
            {
                return null;
            }

            var card = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12, Margin = new Thickness(8) };
            card.Children.Add(BoardTableEditor.ColourSide(ColourChangedFrom, change.From));
            card.Children.Add(BoardTableEditor.ColourSide(ColourChangedTo, change.To));

            return card;
        }

        private static StackPanel ColourSide(string heading, string value)
        {
            var side = new StackPanel { Spacing = 4 };

            side.Children.Add(new TextBlock { Text = heading, FontWeight = FontWeight.SemiBold });
            side.Children.Add(new Border
            {
                Width = 96,
                Height = 48,
                BorderThickness = new Thickness(1),
                BorderBrush = Brushes.DarkGray,
                Background = new SolidColorBrush(RectGeometry.ParseColorOrDefault(value, Colors.Transparent)),
            });
            side.Children.Add(new TextBlock { Text = value.Length > 0 ? value : ColourNotSet });

            return side;
        }
    }
}
