using Avalonia;
using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A HINT FOLLOWING THE POINTER STANDS (owner request, 2026-11-20: "have the hover/helper text follow the mouse, as it now
    // shows below image"; BaselineCaptureWindow's "Click to see it larger"): below and right of the pointer, clear of it - never
    // under it - and kept inside the room it is drawn in: past the right edge it moves left, past the bottom it goes above the
    // pointer. It is drawn in the window and takes no click, unlike a tooltip, whose window under the pointer took the next click
    // (ToolTipPlacement).
    //
    // PURE: the pointer, the hint's size and the room in; the hint's top left corner out.
    // ###########################################################################################
    internal static class PointerHintGeometry
    {
        // How far right of and below the pointer the hint's corner stands - past the arrow of a pointer of the usual size.
        internal const double Right = 16;
        internal const double Below = 20;

        // Kept this far inside the room's edges.
        internal const double Edge = 4;

        internal static Point Place(Point pointer, Size hint, Size room)
        {
            double left = pointer.X + Right;

            if (left + hint.Width > room.Width - Edge)
            {
                left = room.Width - Edge - hint.Width;
            }

            double top = pointer.Y + Below;

            if (top + hint.Height > room.Height - Edge)
            {
                top = pointer.Y - Edge - hint.Height;
            }

            return new Point(Math.Max(0, left), Math.Max(0, top));
        }
    }
}
