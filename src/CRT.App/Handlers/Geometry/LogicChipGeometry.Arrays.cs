using Avalonia;
using Handlers.DataHandling;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // THE CHIPS WHOSE SIGNALS CROSS THE CASE (owner request, 2026-11-16: "OK, now go ahead and make the rest") - the 74240
    // and 74244's buffers and the 74245's transceivers, each from a pin on one side to a pin one row up or down on the
    // other. They stand in ONE COLUMN DOWN THE MIDDLE, each halfway between its two pins' rows, so a wire comes in from its
    // pin, steps half a row to it, and leaves it to step the other half to its pin - the columns' sides are where every step
    // is, the column's inside where nothing but the units is.
    //
    // - THE 240 AND 244: eight buffers, those of one group pointing right and the other's left, by turns. Each group's G
    //   comes in on each of its buffers' top edge from a badge beside it - the inside of the column is free for the badges,
    //   and the groups' buffers stand by turns, so no wire could reach them all.
    // - THE 245: eight transceivers one under the other, and over them the COMMON CONTROL SECTION with DIR and the enable,
    //   as the datasheets' logic symbol draws an array.
    // ###########################################################################################
    public static partial class LogicChipGeometry
    {
        // From the case's edge to a step; a buffer's size; a transceiver's width and its step's distance from it.
        private const double ArrayRun = 14;
        private const double ArrayStep = 6;
        private const double ArrayBufferWidth = 16;
        private const double ArrayBufferHeight = 12;
        private const double ArrayBufferEnableAt = 4;
        private const double ArrayEnableTurn = 14;
        private const double TransceiverRun = 16;
        private const double TransceiverStep = 10;
        private const double TransceiverWidth = 30;

        // How far below the case's notch a symbol in the middle of its top starts.
        private const double NotchClear = 5;

        // ###########################################################################################
        // The 240 and 244: each buffer worked out as if it pointed right - its input from the left - and mirrored when it
        // points left. Its badge stands in the room beside its back, between the wire coming into it and the one leaving
        // the buffer above.
        // ###########################################################################################
        private static LogicInside BuffersFor(LogicChip chip, double nameWidth)
        {
            bool inverted = chip.Units[0].Inverted;
            double buffer = LogicChipGeometry.ArrayBufferWidth + (inverted ? LogicChipGeometry.Bubble : 0);
            double zone = LogicChipGeometry.BadgeWidth + (2 * BadgeClear);
            PinoutLayout layout = LogicChipGeometry.Case(
                chip.PinCount, (2 * (LogicChipGeometry.ArrayRun + LogicChipGeometry.ArrayStep + zone)) + buffer, nameWidth);
            Rect body = layout.Body;
            double back = body.Center.X - (buffer / 2);
            double stepIn = body.Left + LogicChipGeometry.ArrayRun;
            double stepOut = body.Right - LogicChipGeometry.ArrayRun;
            var drawing = new Drawing();

            foreach (LogicUnit unit in chip.Units)
            {
                int input = unit.Inputs[0];
                int output = unit.Output;
                LogicTerminal enable = unit.Of(TerminalRole.Enable).Single();
                double yIn = LogicChipGeometry.YOf(layout, input);
                double yOut = LogicChipGeometry.YOf(layout, output);
                double y = (yIn + yOut) / 2;
                var part = new Drawing();

                // The enable near its back on its top edge, up from its bubble and back to its badge.
                double half = LogicChipGeometry.ArrayBufferHeight / 2;
                double x = back + LogicChipGeometry.ArrayBufferEnableAt;
                var edgePoint = new Point(x, y - half + (LogicChipGeometry.ArrayBufferEnableAt / LogicChipGeometry.ArrayBufferWidth * half));
                var terminal = new Point(x, edgePoint.Y - (enable.ActiveLow ? LogicChipGeometry.Bubble : 0));
                var stub = new Point(back - BadgeClear, y - LogicChipGeometry.ArrayEnableTurn);

                LogicGateShape shape = LogicChipGeometry.GateAt(
                    GateSymbol.Buffer, inverted, chip.Outputs, back, y, LogicChipGeometry.ArrayBufferHeight, LogicChipGeometry.ArrayBufferWidth,
                    [(input, y, string.Empty)], output, [new LogicGateControl(terminal, edgePoint, enable.ActiveLow)], [new LogicTerminalAt(enable.Pin, terminal)], unit);

                part.Symbols.Add(shape);
                part.Wire(input, new Point(body.Left, yIn), new Point(stepIn, yIn), new Point(stepIn, y), new Point(back, y));
                part.Wire(output, shape.Output, new Point(stepOut, y), new Point(stepOut, yOut), new Point(body.Right, yOut));
                part.Wire(enable.Pin, stub, new Point(x, stub.Y), terminal);
                part.Badges.Add(new LogicPinBadge(enable.Pin, LogicChipGeometry.BadgeEndingAt(stub), stub));

                drawing.Add(part, LogicChipGeometry.IsLeft(chip, input) ? null : body.Left + body.Right);
            }

            return drawing.In(layout);
        }

        // ###########################################################################################
        // The 245: transceiver n between An, on the left, and Bn a row lower on the right, its box over the rows between them
        // - so the boxes stand one on the next - and the common control section over the first: DIR straight in from its pin
        // beside it, the enable stepping up from its pin to the section's other side.
        // ###########################################################################################
        private static LogicInside TransceiversFor(LogicChip chip, double nameWidth)
        {
            PinoutLayout layout = LogicChipGeometry.Case(
                chip.PinCount, (2 * (LogicChipGeometry.TransceiverRun + LogicChipGeometry.TransceiverStep)) + LogicChipGeometry.TransceiverWidth, nameWidth);
            Rect body = layout.Body;
            double left = body.Center.X - (LogicChipGeometry.TransceiverWidth / 2);
            double right = left + LogicChipGeometry.TransceiverWidth;
            double stepLeft = left - LogicChipGeometry.TransceiverStep;
            double stepRight = right + LogicChipGeometry.TransceiverStep;
            var drawing = new Drawing();
            double first = double.MaxValue;

            foreach (LogicUnit unit in chip.Units)
            {
                int a = unit.Of(TerminalRole.Input).Single(terminal => terminal.Index == 0).Pin;
                int b = unit.Of(TerminalRole.Input).Single(terminal => terminal.Index == 1).Pin;
                double yA = LogicChipGeometry.YOf(layout, a);
                double yB = LogicChipGeometry.YOf(layout, b);
                double y = (yA + yB) / 2;
                var rect = new Rect(left, System.Math.Min(yA, yB), LogicChipGeometry.TransceiverWidth, System.Math.Abs(yB - yA));
                first = System.Math.Min(first, rect.Top);

                LogicBlockPort portA = LogicChipGeometry.PortOf(a, rect, BlockEdge.Left, y, string.Empty);
                LogicBlockPort portB = LogicChipGeometry.PortOf(b, rect, BlockEdge.Right, y, string.Empty);

                drawing.Wire(a, new Point(body.Left, yA), new Point(stepLeft, yA), new Point(stepLeft, y), portA.At);
                drawing.Wire(b, portB.At, new Point(stepRight, y), new Point(stepRight, yB), new Point(body.Right, yB));
                drawing.Symbols.Add(new LogicBlockShape(rect, [unit], [portA, portB], BlockGlyph.Transceiver));
            }

            // The section over them, with what every transceiver shares - under the case's notch; a control stepping up to it
            // comes in under one beside it, the two marks one over the other.
            double sectionTop = body.Top + LogicChipGeometry.NotchRadius + LogicChipGeometry.NotchClear;
            var section = new Rect(left, sectionTop, LogicChipGeometry.TransceiverWidth, first - sectionTop);
            LogicUnit any = chip.Units[0];
            var ports = new List<LogicBlockPort>();
            double sectionY = section.Bottom - 12;

            foreach (LogicTerminal control in any.Terminals.Where(terminal => terminal.Role is TerminalRole.Direction or TerminalRole.Enable))
            {
                bool pinLeft = LogicChipGeometry.IsLeft(chip, control.Pin);
                double y = LogicChipGeometry.YOf(layout, control.Pin);
                bool straight = y > section.Top && y < sectionY - 4;
                LogicBlockPort port = LogicChipGeometry.PortFor(chip, any, control, section, pinLeft ? BlockEdge.Left : BlockEdge.Right, straight ? y : sectionY);
                ports.Add(port);

                double edge = pinLeft ? body.Left : body.Right;
                double step = pinLeft ? stepLeft : stepRight;

                if (straight)
                {
                    drawing.Wire(control.Pin, new Point(edge, y), port.At);
                }
                else
                {
                    drawing.Wire(control.Pin, new Point(edge, y), new Point(step, y), new Point(step, port.At.Y), port.At);
                }
            }

            drawing.Symbols.Add(new LogicBlockShape(section, chip.Units, ports));

            return drawing.In(layout);
        }
    }
}
