using OfficeOpenXml;
using System;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE TEXT OF A WORKBOOK CELL, THE SAME ON EVERY MACHINE (code review, 2026-10-13).
    //
    // EPPlus's ExcelRange.Text formats a NUMBER (or a date) with the CURRENT culture. The board
    // workbooks are hand-made in Excel, and 13 of the 14 shipped v3 workbooks store a highlight
    // opacity as a number (<v>0.3</v>), so a comma-decimal machine read "0,3" where the server -
    // which runs with the invariant culture - reads "0.3". A draft then wrote "0,3" back as TEXT,
    // the maintainer saw every schematic row changed, and an approval published a value no machine
    // could parse (RectGeometry.ParseOpacityOrDefault is invariant).
    //
    // So every reader of a workbook cell goes through here, and a cell that is not already a
    // string is formatted under the INVARIANT culture - exactly what the server reads - with its
    // own number format still applied ("0.00", "0%"). A string cell is returned as EPPlus gives it,
    // unchanged from before. The caller's culture is put back however the read ends.
    //
    // BoardWorkbookWriter is the other half: it writes every cell as text, so a workbook CRT or
    // the server wrote never holds a number in the first place.
    // ###########################################################################################
    internal static class WorkbookCellText
    {
        public static string Of(ExcelRange cell)
        {
            ArgumentNullException.ThrowIfNull(cell);

            object? value = cell.Value;

            if (value is null)
                return string.Empty;

            CultureInfo caller = CultureInfo.CurrentCulture;

            if (value is string || ReferenceEquals(caller, CultureInfo.InvariantCulture))
                return cell.Text ?? string.Empty;

            try
            {
                CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
                return cell.Text ?? string.Empty;
            }
            finally
            {
                CultureInfo.CurrentCulture = caller;
            }
        }

        // The trimmed text of one cell of a sheet - what every board and master reader wants.
        public static string Trimmed(ExcelWorksheet sheet, int row, int column) =>
            WorkbookCellText.Of(sheet.Cells[row, column]).Trim();
    }
}
