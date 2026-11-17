using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // One of a 4066's switches as drawn: the contact its blade turns on (Pivot) and the one it closes onto (Contact), each
    // where its pin's wire ends, the blade's open tip, the end of its control's wire, and the end of the dashed link from
    // there to the blade.
    // ###########################################################################################
    public sealed record LogicSwitchShape(
        LogicUnit Unit,
        Point Pivot,
        Point Contact,
        Point BladeTip,
        Point Control,
        Point Link,
        Rect Bounds,
        IReadOnlyList<LogicTerminalAt> Terminals) : ILogicSymbol
    {
        public IReadOnlyList<LogicUnit> Units => [this.Unit];
    }

    // ###########################################################################################
    // THE 4066'S SWITCHES (owner request, 2026-11-16: "OK, now go ahead and make the rest"), as its datasheet's functional
    // diagram draws them: each a switch between its two pins, side by side on one side of the case - their wires straight
    // in to its two contacts, one under the other, its blade drawn open from the lower - and its control's wire coming to
    // a dashed link to the blade, from its pin wherever that is: above or below on the same side, or straight across from
    // the other side, over the rows no other switch uses there.
    // ###########################################################################################
    public static partial class LogicChipGeometry
    {
        // From the case's edge to the contacts; from them to the control's end and on to where its wire turns.
        private const double ContactRun = 26;
        private const double ControlAt = 22;
        private const double ControlTurn = 8;

        private static LogicInside SwitchesFor(LogicChip chip, double nameWidth)
        {
            double reach = LogicChipGeometry.ContactRun + LogicChipGeometry.ControlAt + LogicChipGeometry.ControlTurn + 6;
            PinoutLayout layout = LogicChipGeometry.Case(chip.PinCount, (2 * reach) + LogicChipGeometry.CentreGap, nameWidth);
            var drawing = new Drawing();

            foreach (LogicUnit unit in chip.Units)
            {
                IReadOnlyList<int> sides = unit.Inputs;

                if (LogicChipGeometry.IsLeft(chip, sides[0]) != LogicChipGeometry.IsLeft(chip, sides[1]) ||
                    Math.Abs(LogicChipGeometry.RowOf(chip, sides[0]) - LogicChipGeometry.RowOf(chip, sides[1])) != 1)
                {
                    throw new InvalidOperationException($"{chip.Name}: a switch's two pins must be next to each other on one side.");
                }

                bool left = LogicChipGeometry.IsLeft(chip, sides[0]);
                var part = new Drawing();
                LogicChipGeometry.SwitchOnLeft(chip, unit, left, layout, part);
                drawing.Add(part, left ? null : layout.Body.Left + layout.Body.Right);
            }

            return drawing.In(layout);
        }

        // A switch worked out as if it stood on the left.
        private static void SwitchOnLeft(LogicChip chip, LogicUnit unit, bool left, PinoutLayout layout, Drawing part)
        {
            Rect body = layout.Body;
            double Y(int pin) => LogicChipGeometry.YOf(layout, pin);

            int[] sides = [.. unit.Inputs.OrderBy(pin => Y(pin))];
            int contactPin = sides[0];
            int pivotPin = sides[1];
            int controlPin = unit.Of(TerminalRole.Control).Single().Pin;

            double x = body.Left + LogicChipGeometry.ContactRun;
            var contact = new Point(x, Y(contactPin));
            var pivot = new Point(x, Y(pivotPin));
            var tip = new Point(x + 10, contact.Y + 7);
            var link = new Point((pivot.X + tip.X) / 2, (pivot.Y + tip.Y) / 2);
            var control = new Point(x + LogicChipGeometry.ControlAt, link.Y);
            double turn = control.X + LogicChipGeometry.ControlTurn;

            part.Wire(contactPin, new Point(body.Left, contact.Y), contact);
            part.Wire(pivotPin, new Point(body.Left, pivot.Y), pivot);

            // The control: from its pin - this side or across from the other - turning to the height of the link.
            double controlY = Y(controlPin);
            double from = LogicChipGeometry.IsLeft(chip, controlPin) == left ? body.Left : body.Right;
            part.Wire(controlPin, new Point(from, controlY), new Point(turn, controlY), new Point(turn, control.Y), control);

            var bounds = new Rect(new Point(x - 4, contact.Y - 4), new Point(control.X, pivot.Y + 4));

            part.Symbols.Add(new LogicSwitchShape(
                unit,
                pivot,
                contact,
                tip,
                control,
                link,
                bounds,
                [new LogicTerminalAt(contactPin, contact), new LogicTerminalAt(pivotPin, pivot), new LogicTerminalAt(controlPin, control)]));
        }
    }
}
