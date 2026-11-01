using Avalonia;
using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A DIALOG THAT SIZES ITSELF TO ITS CONTENT GOES (owner report, 2026-11-01, of the Submit
    // dialog: "it seems to take full height, and where it has no scrollbars, so I cannot see the submit
    // button ... the next popup ... is shown in the top - I think it should be shown in the middle").
    //
    //   - NO TALLER THAN THE SCREEN IT IS ON (MaxHeight): its working area - the screen less the task
    //     bar - with a margin above and below, so its buttons are always on screen and its content
    //     scrolls instead.
    //   - CENTRED OVER ITS OWNER (CentredOver) whenever its size changes: a window sized to its content
    //     keeps its top edge where it was as it shrinks, so the dialog's next, smaller panel sat at the
    //     top of where the tall one had been. Kept inside the working area.
    //
    // A window's position and the screen's working area are in DEVICE pixels; its size in device-
    // independent units, times `scaling` on the screen.
    // ###########################################################################################
    public static class DialogPlacement
    {
        // The room kept free above and below a dialog on its screen, in device-independent units.
        public const double ScreenMargin = 24;

        // A dialog is never squeezed below this, however small the screen.
        public const double LowestMaxHeight = 240;

        // ###########################################################################################
        // The tallest a dialog may be on a screen whose working area is `workingAreaHeight` device
        // pixels high at `scaling` - in device-independent units.
        // ###########################################################################################
        public static double MaxHeight(int workingAreaHeight, double scaling)
        {
            double height = workingAreaHeight / (scaling > 0 ? scaling : 1.0);

            return Math.Max(LowestMaxHeight, height - (2 * ScreenMargin));
        }

        // ###########################################################################################
        // The top-left corner, in device pixels, of a dialog `size` big (its frame, device-independent
        // units) centred over `owner` (device pixels) - moved inside `workingArea` where it would stand
        // out of it, its top edge kept on screen first.
        // ###########################################################################################
        public static PixelPoint CentredOver(PixelRect owner, Size size, double scaling, PixelRect workingArea)
        {
            double scale = scaling > 0 ? scaling : 1.0;
            int width = (int)Math.Round(size.Width * scale);
            int height = (int)Math.Round(size.Height * scale);

            int x = owner.X + ((owner.Width - width) / 2);
            int y = owner.Y + ((owner.Height - height) / 2);

            if (workingArea.Width > 0 && workingArea.Height > 0)
            {
                x = Math.Max(workingArea.X, Math.Min(x, workingArea.Right - width));
                y = Math.Max(workingArea.Y, Math.Min(y, workingArea.Bottom - height));
            }

            return new PixelPoint(x, y);
        }
    }
}
