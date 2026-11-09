using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A BOARD'S REGION BUTTONS STAND (2026-11-09: a board may come in regions of its own) - under the component list
    // and on the component popup: side by side as wide as they can be, at most PerRow to a row, the next row under it.
    // Two regions - PAL and NTSC, every published board - are the two halves they always were.
    //
    // *** THE GRIDS NEVER CHANGE THEIR ROWS OR COLUMNS *** (owner report, 2026-11-09: "Launching CRT app now gives me a stack
    // trace directly"). Avalonia 12.1.4's Grid keeps each child's cell as last laid out and does not look again when its
    // definitions change - only a child, or its Grid.Row/Column, changing makes it - so a grid that lost a row or column
    // measured a button against one it no longer had and threw (ArgumentOutOfRangeException in GetMeasureSizeForRange). So
    // the list's grid is GridColumns equal columns 6px apart, a button spanning GridColumns / its row's count of them - two
    // span three each, exactly the halves "*,6,*" gave, and three span two - over two rows, the second 6px lower by its
    // buttons' margin; the popup's is two Auto columns, the second's buttons 6px to the right by theirs (an empty Auto column
    // is no width, and a margin on no button no gap). CLAUDE.md: buttons side by side are 6px apart.
    // ###########################################################################################
    internal static class RegionButtonLayout
    {
        public const int PerRow = 3;

        // The list's grid: six equal columns (two, three and one to a row all fit them) 6px apart, and two rows.
        public const int GridColumns = 6;

        public const double Gap = 6;

        // How many buttons stand in a row, and how many rows there are, for `count` buttons - four stand two by two, as one
        // alone on a row of its own would be a third as wide as the rest (the render, 2026-11-09).
        public static (int Columns, int Rows) Of(int count)
        {
            if (count <= 0)
            {
                return (0, 0);
            }

            int columns = count == 4 ? 2 : Math.Min(count, RegionButtonLayout.PerRow);

            return (columns, (count + columns - 1) / columns);
        }

        // ###########################################################################################
        // Where button `index` of `count` stands in the list's fixed grid (GridColumns columns, two rows): its row, its first
        // column and how many columns it spans.
        // ###########################################################################################
        public static (int Row, int Column, int ColumnSpan) PlaceOf(int index, int count)
        {
            (int columns, _) = RegionButtonLayout.Of(count);

            if (columns == 0)
            {
                return (0, 0, RegionButtonLayout.GridColumns);
            }

            int span = RegionButtonLayout.GridColumns / columns;

            return (index / columns, (index % columns) * span, span);
        }

        // ###########################################################################################
        // The component popup stacks them instead, in the three rows of its footer, a column of PerRow, the next column
        // beside it - EACH COLUMN FROM THE BOTTOM: fewer than three in it stand at the bottom - PAL and NTSC in the second
        // and third rows, as they always did, and a fourth region beside the third, not at the top beside the first (code
        // review, 2026-11-09). The cell in its region grid (rows 28,4,28,4,28; two Auto columns, the second's buttons 6px in
        // by their margin).
        // ###########################################################################################
        public static (int Row, int Column) PopupCellOf(int index, int count)
        {
            if (count <= 0)
            {
                return (0, 0);
            }

            int column = index / RegionButtonLayout.PerRow;
            int inColumn = Math.Min(RegionButtonLayout.PerRow, count - (column * RegionButtonLayout.PerRow));
            int row = index % RegionButtonLayout.PerRow;

            return ((RegionButtonLayout.PerRow - inColumn + row) * 2, column);
        }
    }
}
