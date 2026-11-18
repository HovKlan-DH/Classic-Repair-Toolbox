using Avalonia;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // A PIN'S NUMBER ON ITS PAD (PinNumbersOverlay - owner request, 2026-11-17: "a pin number badge for all component pins ...
    // placed directly on top of the pin"): as tall as the pad is on screen, so neighbours stay apart as far as the board lets
    // them, but never below what a number can be read at (LeastHeight) nor larger than a badge needs to be (MostHeight);
    // centered on the pad, as wide as its number with a little room each side, never narrower than it is tall.
    //
    // PURE: sizes in, sizes and a rectangle out.
    // ###########################################################################################
    internal static class PinBadgeGeometry
    {
        // A little larger since 2026-11-17 (owner request: "make the pin text a little larger, as it is hard to see. Maybe make
        // the box size a little larger also" - 12 to 18 tall, the number 0.68 of it, until then).
        public const double LeastHeight = 14;
        public const double MostHeight = 20;

        // The room each side of a number.
        public const double SideRoom = 3;

        public static double HeightFor(double padOnScreen) =>
            double.IsFinite(padOnScreen) ? Math.Clamp(padOnScreen, LeastHeight, MostHeight) : LeastHeight;

        // The number's size in a badge `height` tall.
        public static double TextSizeFor(double height) => Math.Round(height * 0.74, 1);

        public static Rect AreaOf(Point center, double height, double numberWidth)
        {
            double width = Math.Max(height, numberWidth + (2 * SideRoom));

            return new Rect(center.X - (width / 2), center.Y - (height / 2), width, height);
        }

        // ###########################################################################################
        // PIN 1'S ORANGE PAD FRAMED IN BLACK, not numbered (PinNumbersOverlay - owner request, 2026-11-17): the pad of
        // `padSize` in the image's coordinates, centred on `centre` on screen, at `scale` - its orange ending half the
        // overlay's outline outside its size (KiCadPadGeometry.OutlineWidth, at that scale) - and the frame a 1-wide line
        // just outside the orange, half a pixel further. Here, not in the control, so it is tested and follows the overlay's
        // outline (code review, 2026-11-18: its 0.6 was a copy). Turned about `centre` by the pad's own turn as it is drawn.
        // ###########################################################################################
        public static Rect FirstPinFrame(Point centre, Size padSize, double scale)
        {
            double width = padSize.Width * scale;
            double height = padSize.Height * scale;

            return new Rect(centre.X - (width / 2), centre.Y - (height / 2), width, height)
                .Inflate(KiCadPadGeometry.OutlineWidth / 2 * scale)
                .Inflate(0.5);
        }

        // ###########################################################################################
        // Which of a footprint's pad numbers is its first pin - told apart among the numbers: "1" where it has one, else the
        // first by KiCad's order of designators ("A1" before "A2", "2" before "10"), as pin 1's orange pad is chosen.
        // "" for none.
        // ###########################################################################################
        public static string FirstPinOf(IEnumerable<string> numbers)
        {
            ArgumentNullException.ThrowIfNull(numbers);

            List<string> named = numbers.Select(number => number?.Trim() ?? string.Empty).Where(number => number.Length > 0).ToList();

            return named.Contains("1")
                ? "1"
                : named.OrderBy(number => number, Comparer<string>.Create(KiCadLayerGeometry.ComparePadDesignators)).FirstOrDefault() ?? string.Empty;
        }
    }
}
