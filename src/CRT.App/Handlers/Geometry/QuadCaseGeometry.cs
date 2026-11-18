using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // A CHIP WITH PINS ON FOUR SIDES, DRAWN SQUARE (owner request, 2026-11-18: "Drawing square packages ... yes, I want ...
    // because there might come new hardware soon - e.g. Amiga 500 or Amiga 1000", and "It has a lot of SMD components"): a
    // PLCC or a QFP seen from above (ComponentVocabulary.ChipCaseOf), in the very style of the DIP's case
    // (LogicChipGeometry.Case) - the body, each pin a square beside it with its number, a lead joining them, its name outside
    // - so PinoutDrawing paints it, colours its pins by direction and finds the pin under the pointer as on any case.
    //
    // THE PINS RUN COUNTER-CLOCKWISE, AS EVERY MAKER NUMBERS THEM: a QFP's pin 1 at the top of the left side, down it, along
    // the bottom, up the right and back along the top; a PLCC's pin 1 in the MIDDLE of the top side, then the same way round -
    // along the top to its left corner, down the left, along the bottom, up the right and along the top back to pin 2's
    // other neighbour. The cut corner is the top left's, the dot by pin 1 inside the body.
    //
    // A NAME ABOVE OR UNDER THE CASE IS WRITTEN UPRIGHT ON ITS SIDE, reading from the bottom up, as the datasheets write
    // them: its foot at its pin on the top side, its head at its pin on the bottom side. So the room for names - nameWidth,
    // the longest name's (PinoutGeometry.NameWidthFor) - stands on every side, and the corners stay empty.
    //
    // Three-digit pin numbers (100 pins and more) take a wider square and pitch. Coordinates start at (0, 0).
    // ###########################################################################################
    public static class QuadCaseGeometry
    {
        // The cut corner - the top left's - and how far into the body the dot by pin 1 stands.
        public const double Bevel = 10;
        public const double DotInset = 12;

        // From a square's outer edge to its name, as on the DIP's case.
        public const double NameGap = 4;

        public static double SquareFor(int pins) => pins >= 100 ? 22 : LogicChipGeometry.Square;

        public static double PitchFor(int pins) => pins >= 100 ? 26 : LogicChipGeometry.Pitch;

        // ###########################################################################################
        // The case of `chipCase` (a PLCC or a QFP), its names `nameWidth` wide on every side. Its pins in number order.
        // ###########################################################################################
        public static PinoutLayout Case(ChipCase chipCase, double nameWidth = PinoutGeometry.NameWidth)
        {
            ArgumentNullException.ThrowIfNull(chipCase);

            if (chipCase.Kind == ChipCaseKind.TwoRows || chipCase.Pins % 4 != 0 || chipCase.Pins < 8)
            {
                throw new ArgumentOutOfRangeException(nameof(chipCase), chipCase, "A case of four sides has a quarter of its pins on each.");
            }

            int perSide = chipCase.PerSide;
            double pitch = QuadCaseGeometry.PitchFor(chipCase.Pins);
            double square = QuadCaseGeometry.SquareFor(chipCase.Pins);
            double gap = LogicChipGeometry.LeadGap;
            double padding = LogicChipGeometry.EndPadding;

            double margin = nameWidth + square + gap;
            double side = (perSide * pitch) + (2 * padding);
            var body = new Rect(margin, margin, side, side);

            // The ring of places counter-clockwise from the top of the left side, and where pin 1 stands in it.
            int start = chipCase.Kind == ChipCaseKind.Plcc ? (4 * perSide) - ((perSide + 1) / 2) : 0;
            var pins = new PinoutPin[chipCase.Pins];

            for (int place = 0; place < chipCase.Pins; place++)
            {
                int number = ((place - start + chipCase.Pins) % chipCase.Pins) + 1;
                int index = place % perSide;
                double along = padding + ((index + 0.5) * pitch);

                pins[number - 1] = (place / perSide) switch
                {
                    // Down the left side.
                    0 => QuadCaseGeometry.Pin(number, PinSide.Left, new Point(body.Left - gap - (square / 2), body.Top + along), square),

                    // Along the bottom, left to right.
                    1 => QuadCaseGeometry.Pin(number, PinSide.Bottom, new Point(body.Left + along, body.Bottom + gap + (square / 2)), square),

                    // Up the right side.
                    2 => QuadCaseGeometry.Pin(number, PinSide.Right, new Point(body.Right + gap + (square / 2), body.Bottom - along), square),

                    // Along the top, right to left.
                    _ => QuadCaseGeometry.Pin(number, PinSide.Top, new Point(body.Right - along, body.Top - gap - (square / 2)), square),
                };
            }

            var size = new Size(body.Right + gap + square + nameWidth, body.Bottom + gap + square + nameWidth);

            return new PinoutLayout(body, pins, size, pitch);
        }

        // ###########################################################################################
        // The dot by pin 1, inside the body: in from pin 1's side, across from its square.
        // ###########################################################################################
        public static Point FirstPinDotOf(PinoutLayout layout)
        {
            ArgumentNullException.ThrowIfNull(layout);

            PinoutPin first = layout.Pins.First(pin => pin.Number == 1);
            Rect body = layout.Body;

            return first.Side switch
            {
                PinSide.Left => new Point(body.Left + QuadCaseGeometry.DotInset, first.Lead.Center.Y),
                PinSide.Right => new Point(body.Right - QuadCaseGeometry.DotInset, first.Lead.Center.Y),
                PinSide.Bottom => new Point(first.Lead.Center.X, body.Bottom - QuadCaseGeometry.DotInset),
                _ => new Point(first.Lead.Center.X, body.Top + QuadCaseGeometry.DotInset),
            };
        }

        // The body's outline: a square with its top left corner cut.
        public static IReadOnlyList<Point> OutlineOf(Rect body) =>
        [
            new Point(body.Left + QuadCaseGeometry.Bevel, body.Top),
            body.TopRight,
            body.BottomRight,
            body.BottomLeft,
            new Point(body.Left, body.Top + QuadCaseGeometry.Bevel),
        ];

        // A pin whose square stands at `centre`: its name's anchor beyond the square's outer edge.
        private static PinoutPin Pin(int number, PinSide side, Point centre, double square)
        {
            var lead = new Rect(centre.X - (square / 2), centre.Y - (square / 2), square, square);

            Point name = side switch
            {
                PinSide.Left => new Point(lead.Left - QuadCaseGeometry.NameGap, centre.Y),
                PinSide.Right => new Point(lead.Right + QuadCaseGeometry.NameGap, centre.Y),
                PinSide.Bottom => new Point(centre.X, lead.Bottom + QuadCaseGeometry.NameGap),
                _ => new Point(centre.X, lead.Top - QuadCaseGeometry.NameGap),
            };

            return new PinoutPin(number, side, lead, centre, name);
        }
    }
}
