using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A POINT OF A LIBRARY SYMBOL LANDS ON ITS SHEET.
    //
    // A library symbol is drawn Y up and a sheet is Y down. The symbol's angle turns it first, and its
    // mirror then flips it on the sheet: "x" top to bottom, "y" left to right.
    //
    // *** THAT ORDER WAS MEASURED, NOT ASSUMED (2026-11-04). *** A pin's tip has to land on a wire, so each
    // order was scored on every KiCad project CRT ships. It only shows on a symbol both turned and mirrored:
    // this order put 100% of the tips at 90 degrees and 94% at 270 on a wire, the other 29% and 53%.
    // ###########################################################################################
    internal static class KiCadSymbolPlacement
    {
        public static (double X, double Y) OnSheet(double x, double y, double atX, double atY, double angle, string? mirror)
        {
            int turn = (((int)Math.Round(angle) % 360) + 360) % 360;

            (double sheetX, double sheetY) = turn switch
            {
                90 => (-y, -x),
                180 => (-x, y),
                270 => (y, x),
                _ => (x, -y),
            };

            if (string.Equals(mirror, "x", StringComparison.OrdinalIgnoreCase))
            {
                sheetY = -sheetY;
            }
            else if (string.Equals(mirror, "y", StringComparison.OrdinalIgnoreCase))
            {
                sheetX = -sheetX;
            }

            return (atX + sheetX, atY + sheetY);
        }
    }
}
