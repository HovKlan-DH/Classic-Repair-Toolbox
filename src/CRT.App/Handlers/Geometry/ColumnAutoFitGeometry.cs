using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // Which column a press on a column heading's EDGE belongs to - see ColumnAutoFitGeometry.EdgeAt.
    internal enum ColumnEdge
    {
        None,
        ThisColumn,
        PreviousColumn,
    }

    // ###########################################################################################
    // DOUBLE-CLICKING A COLUMN HEADING'S EDGE FITS THE COLUMN TO ITS TEXT, as in Excel (owner request,
    // 2026-10-04: "doubleclicking the column resizer should auto-fit column to show all text - just
    // like Excel does"). The board table measures (BoardTableEditor.TextWrap.cs), and this decides.
    //
    // *** EVERY ROW, NOT THE ROWS ON SCREEN. *** The grid has a double-click fit of its own
    // (ProDataGrid's CanUserResizeColumnsOnDoubleClick), but it measures only the rows it has built
    // - the ones in view - so a long note further down was left out and the column fitted nothing
    // the user had asked about. The table turns that off and measures every row it shows.
    //
    // *** NEVER WIDER THAN THE TABLE CAN SHOW. *** Excel makes a column as wide as its longest text
    // even when that runs past the screen. The table's cells always wrap, so a column as wide as the
    // table already shows all of a long text, over several lines - one wider would only hide the end
    // of each line past the right edge.
    // ###########################################################################################
    internal static class ColumnAutoFitGeometry
    {
        // How close to a heading's edge a press counts as on the edge - ProDataGrid's own resize
        // grip (DataGridColumnHeader: 5px either side), so the fit answers exactly where the
        // resize cursor shows.
        public const double EdgeGrip = 5;

        // Added to the widest text: a column exactly as wide as a text can still wrap its last word
        // once the layout rounds to whole pixels.
        public const double RoundingSlack = 1;

        // ###########################################################################################
        // Which column a press at x (in the heading's own coordinates) resizes: its right edge is
        // this column's, its left edge the previous column's - the right edge first, as the grid
        // decides it, so a heading narrower than both grips answers for itself.
        // ###########################################################################################
        public static ColumnEdge EdgeAt(double x, double headerWidth)
        {
            if (headerWidth <= 0 || double.IsNaN(x))
            {
                return ColumnEdge.None;
            }

            if (headerWidth - x <= ColumnAutoFitGeometry.EdgeGrip)
            {
                return ColumnEdge.ThisColumn;
            }

            if (x <= ColumnAutoFitGeometry.EdgeGrip)
            {
                return ColumnEdge.PreviousColumn;
            }

            return ColumnEdge.None;
        }

        // ###########################################################################################
        // The width that shows the heading and every text on one line - each text plus the cell's
        // own margins and lines (cellChrome) - within minWidth and maxWidth, in whole pixels. A
        // maxWidth below minWidth gives minWidth: the column's own minimum always wins.
        // ###########################################################################################
        public static double FitWidth(double headerWidth, IEnumerable<double> textWidths, double cellChrome, double minWidth, double maxWidth)
        {
            double widest = double.IsNaN(headerWidth) ? 0 : Math.Max(0, headerWidth);
            double chrome = double.IsNaN(cellChrome) ? 0 : Math.Max(0, cellChrome);

            foreach (double textWidth in textWidths)
            {
                if (double.IsNaN(textWidth) || textWidth <= 0)
                {
                    continue;
                }

                widest = Math.Max(widest, textWidth + chrome + ColumnAutoFitGeometry.RoundingSlack);
            }

            double lowest = double.IsNaN(minWidth) ? 0 : Math.Max(0, minWidth);
            double highest = double.IsNaN(maxWidth) || maxWidth < lowest ? lowest : maxWidth;

            return Math.Clamp(Math.Ceiling(widest), lowest, highest);
        }
    }
}
