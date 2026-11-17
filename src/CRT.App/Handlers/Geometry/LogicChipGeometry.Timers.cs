using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // THE 555'S AND 556'S TIMERS (owner request, 2026-11-16: "OK, now go ahead and make the rest"), as their datasheets'
    // functional diagrams draw one: two comparators - THRES against CONT (two thirds of the supply) and TRIG against a third
    // of it - setting and resetting a flip-flop, which RESET clears, and the flip-flop driving the output stage and,
    // inverted, the open-collector discharge. Its parts are a gate's shapes and a block; the wires between them are
    // inside the timer (pin 0), the rest straight from their pins. The 556's two stand each on its side, their pins in the
    // same order down it; the 555's are on both sides of its case.
    // ###########################################################################################
    public static partial class LogicChipGeometry
    {
        // From the case's edge to an output stage's tip and to a comparator's back - room for the third's mark before it -
        // and between TRIG and the third; from the edge to the flip-flop, and its width.
        private const double StageTip = 14;
        private const double ComparatorBack = 25;
        private const double ReferenceApart = 12;
        private const double TimerFlipFlopAt = 62;
        private const double TimerFlipFlopWidth = 30;
        private const double TimerEndPadding = 20;

        // ###########################################################################################
        // *** THE COMPARATOR'S "+" AND "-" STAND CLEAR OF ITS OUTLINE (owner request, 2026-11-17: "the +/- is too close to
        // its borders"). *** Each is drawn as strokes, not a font's glyph, so where its ink lies is known: SignArm each way
        // from its middle, on its input's line, its back arm MarkInset from the back - and the comparator is as tall as its
        // signs need (ComparatorHalfHeight, LogicChipGeometry.Marks.cs).
        // ###########################################################################################
        public const double SignArm = 2.5;
        public const double SignInset = LogicChipGeometry.MarkInset + LogicChipGeometry.SignArm;

        // ###########################################################################################
        // *** "1/3" AT THE SIZE OF THE OTHER MARKS (owner request, 2026-11-17: 'this "1/3" text could have same text size as
        // the other texts inside this IC (letters in gates)'). *** What the internal reference of a third of the supply is
        // written as, outside its comparator's back at the end of a stub of wire TagLine long, in an area TagWidth wide up to
        // TagGap from the stub - plain ASCII, as every text CRT shows; the "⅓" glyph was half the size of an "R". The area
        // holds "1/3" at the marks' size with room to spare (11.5 wide in Segoe UI), and LogicChipGeometryTests keep it as
        // clear of the case's edge as a badge and off every wire and symbol.
        // ###########################################################################################
        public const string Third = "1/3";
        public const double TagLine = 4;
        public const double TagGap = 1.5;
        public const double TagWidth = 13;
        public const double TagHeight = 11;

        // The area the tag at input `index` is written in: outside the back, beyond its line, level with the input.
        public static Rect TagArea(LogicGateShape gate, int index)
        {
            ArgumentNullException.ThrowIfNull(gate);

            double sign = gate.Facing == GateFacing.Right ? 1 : -1;
            Point input = gate.Inputs[index];
            double near = input.X - (sign * (LogicChipGeometry.TagLine + LogicChipGeometry.TagGap));
            double far = near - (sign * LogicChipGeometry.TagWidth);

            return new Rect(Math.Min(near, far), input.Y - (LogicChipGeometry.TagHeight / 2), LogicChipGeometry.TagWidth, LogicChipGeometry.TagHeight);
        }

        private static LogicInside TimersFor(LogicChip chip, double nameWidth)
        {
            double reach = LogicChipGeometry.TimerFlipFlopAt + LogicChipGeometry.TimerFlipFlopWidth + 6;

            if (chip.Units.Count == 1)
            {
                PinoutLayout single = LogicChipGeometry.Case(chip.PinCount, (2 * LogicChipGeometry.TimerFlipFlopAt) + LogicChipGeometry.TimerFlipFlopWidth, nameWidth);
                var whole = new Drawing();
                LogicChipGeometry.TimerAcross(chip, chip.Units[0], single, whole);

                return whole.In(single);
            }

            // More room at the ends: the 556's second TRIG is its last row, and its comparator's third stands under it.
            PinoutLayout layout = LogicChipGeometry.Case(chip.PinCount, (2 * reach) + LogicChipGeometry.CentreGap, nameWidth, LogicChipGeometry.TimerEndPadding);
            var drawing = new Drawing();

            foreach (LogicUnit unit in chip.Units)
            {
                bool left = LogicChipGeometry.IsLeft(chip, unit.Pins.First());

                if (unit.Pins.Any(pin => LogicChipGeometry.IsLeft(chip, pin) != left))
                {
                    throw new InvalidOperationException($"{chip.Name}: each timer of a dual timer is on one side.");
                }

                var part = new Drawing();
                LogicChipGeometry.TimerOnLeft(chip, unit, left, layout, part);
                drawing.Add(part, left ? null : layout.Body.Left + layout.Body.Right);
            }

            return drawing.In(layout);
        }

        private static int PinOf(LogicUnit unit, TerminalRole role) => unit.Of(role).Single().Pin;

        // ###########################################################################################
        // The comparator of THRES (+) and CONT (-), its back `back` from the edge, pointing right - and its output.
        // ###########################################################################################
        private static LogicGateShape UpperComparator(LogicUnit unit, PinoutLayout layout, double back)
        {
            int threshold = LogicChipGeometry.PinOf(unit, TerminalRole.Threshold);
            int control = LogicChipGeometry.PinOf(unit, TerminalRole.Control);
            double yThreshold = LogicChipGeometry.YOf(layout, threshold);
            double yControl = LogicChipGeometry.YOf(layout, control);

            return LogicChipGeometry.GateAt(
                GateSymbol.Comparator, inverted: false, LogicOutputStage.TotemPole, back, (yThreshold + yControl) / 2,
                2 * LogicChipGeometry.ComparatorHalfHeight(Math.Abs(yControl - yThreshold) / 2), LogicChipGeometry.BodyWidthOf(GateSymbol.Comparator),
                [(threshold, yThreshold, "+"), (control, yControl, "−")], outputPin: 0, [], [], unit);
        }

        // ###########################################################################################
        // The comparator of TRIG (-) and a third of the supply (+): the third on the row above TRIG when no pin there is
        // used, else under it; its back `back` from the edge, pointing right.
        // ###########################################################################################
        private static LogicGateShape LowerComparator(LogicChip chip, LogicUnit unit, PinoutLayout layout, double back, bool left)
        {
            int trigger = LogicChipGeometry.PinOf(unit, TerminalRole.Trigger);
            double yTrigger = LogicChipGeometry.YOf(layout, trigger);
            int row = LogicChipGeometry.RowOf(chip, trigger);
            bool aboveFree = row > 0 && !LogicChipGeometry.UsedPins(chip).Contains(LogicChipGeometry.PinAtRow(chip, row - 1, left));
            double yThird = aboveFree ? yTrigger - LogicChipGeometry.ReferenceApart : yTrigger + LogicChipGeometry.ReferenceApart;

            LogicGateShape shape = LogicChipGeometry.GateAt(
                GateSymbol.Comparator, inverted: false, LogicOutputStage.TotemPole, back, (yTrigger + yThird) / 2,
                2 * LogicChipGeometry.ComparatorHalfHeight(LogicChipGeometry.ReferenceApart / 2), LogicChipGeometry.BodyWidthOf(GateSymbol.Comparator),
                [(trigger, yTrigger, "−"), (0, yThird, "+")], outputPin: 0, [], [], unit);

            // The third's mark stands outside the back, at the input it is - on the row no pin uses, or under the last, so
            // nothing else is there (its bounds stay the triangle's: grown round the mark, they would take in TRIG's wire).
            return shape with { InputTags = [string.Empty, LogicChipGeometry.Third] };
        }

        // ###########################################################################################
        // An output stage pointing back to its pin on the edge, its tip StageTip from it: the output buffer from Q, or the
        // discharge - inverted, open collector - from Q-bar.
        // ###########################################################################################
        private static LogicGateShape OutputStage(LogicUnit unit, PinoutLayout layout, int pin, bool discharge)
        {
            double edge = layout.Body.Left;
            double y = LogicChipGeometry.YOf(layout, pin);
            double width = LogicChipGeometry.BodyWidthOf(GateSymbol.Buffer);
            double whole = width + (discharge ? LogicChipGeometry.Bubble : 0);
            double tip = edge + LogicChipGeometry.StageTip;

            LogicGateShape shape = LogicChipGeometry.GateAt(
                GateSymbol.Buffer, inverted: discharge, discharge ? LogicOutputStage.OpenCollector : LogicOutputStage.TotemPole,
                tip, y, LogicChipGeometry.BufferHeight, width, [(0, y, string.Empty)], pin, [], [], unit);

            // Turned about its own middle, to point at its pin.
            return (LogicGateShape)LogicChipGeometry.Mirrored(shape, (2 * tip) + whole);
        }

        // ###########################################################################################
        // A timer whose pins are all on one side (the 556's), worked out as if on the left: the output stages and comparators
        // by their pins, the flip-flop inward of them with each of its terminals level with what it is joined to - Q-bar with
        // the discharge, R with the upper comparator, RESET's R with its pin, Q with the output stage, S with the lower
        // comparator - so every wire is straight.
        // ###########################################################################################
        private static void TimerOnLeft(LogicChip chip, LogicUnit unit, bool left, PinoutLayout layout, Drawing part)
        {
            double edge = layout.Body.Left;
            double Y(int pin) => LogicChipGeometry.YOf(layout, pin);

            int reset = LogicChipGeometry.PinOf(unit, TerminalRole.Reset);
            int output = LogicChipGeometry.PinOf(unit, TerminalRole.Output);
            int discharge = LogicChipGeometry.PinOf(unit, TerminalRole.Discharge);

            LogicGateShape upper = LogicChipGeometry.UpperComparator(unit, layout, edge + LogicChipGeometry.ComparatorBack);
            LogicGateShape lower = LogicChipGeometry.LowerComparator(chip, unit, layout, edge + LogicChipGeometry.ComparatorBack, left);
            LogicGateShape stage = LogicChipGeometry.OutputStage(unit, layout, output, discharge: false);
            LogicGateShape drain = LogicChipGeometry.OutputStage(unit, layout, discharge, discharge: true);

            double x = edge + LogicChipGeometry.TimerFlipFlopAt;
            double[] heights = [Y(discharge), upper.CentreY, Y(reset), Y(output), lower.CentreY];
            var rect = new Rect(x, heights.Min() - 12, LogicChipGeometry.TimerFlipFlopWidth, heights.Max() - heights.Min() + 24);

            LogicBlockPort qBar = LogicChipGeometry.PortOf(0, rect, BlockEdge.Left, Y(discharge), "Q", overline: true);
            LogicBlockPort r = LogicChipGeometry.PortOf(0, rect, BlockEdge.Left, upper.CentreY, "R");
            LogicBlockPort clear = LogicChipGeometry.PortOf(reset, rect, BlockEdge.Left, Y(reset), "R", bubble: true);
            LogicBlockPort q = LogicChipGeometry.PortOf(0, rect, BlockEdge.Left, Y(output), "Q");
            LogicBlockPort s = LogicChipGeometry.PortOf(0, rect, BlockEdge.Left, lower.CentreY, "S");

            part.Symbols.AddRange([upper, lower, stage, drain, new LogicBlockShape(rect, [unit], [qBar, r, clear, q, s])]);
            LogicChipGeometry.WireInputs([upper, lower], edge, part);

            part.Wire(discharge, new Point(edge, Y(discharge)), drain.Output);
            part.Wire(0, qBar.At, drain.Inputs[0]);
            part.Wire(0, upper.Output, r.At);
            part.Wire(reset, new Point(edge, Y(reset)), clear.At);
            part.Wire(output, new Point(edge, Y(output)), stage.Output);
            part.Wire(0, q.At, stage.Inputs[0]);
            part.Wire(0, lower.Output, s.At);
        }

        // The comparators' inputs wired straight from their pins - the third, inside the timer, has none.
        private static void WireInputs(IEnumerable<LogicGateShape> comparators, double edge, Drawing part)
        {
            foreach (LogicGateShape comparator in comparators)
            {
                foreach (LogicTerminalAt input in comparator.Terminals.Take(comparator.Inputs.Count).Where(terminal => terminal.Pin != 0))
                {
                    part.Wire(input.Pin, new Point(edge, input.At.Y), input.At);
                }
            }
        }

        // ###########################################################################################
        // The 555, its pins on both sides: TRIG's comparator, the output stage and RESET on the left, THRES and CONT's
        // comparator and the discharge on the right - worked out as if on the left and mirrored - and the flip-flop in the
        // middle, each of its terminals level with what it is joined to.
        // ###########################################################################################
        private static void TimerAcross(LogicChip chip, LogicUnit unit, PinoutLayout layout, Drawing drawing)
        {
            Rect body = layout.Body;
            double edge = body.Left;
            double Y(int pin) => LogicChipGeometry.YOf(layout, pin);

            int reset = LogicChipGeometry.PinOf(unit, TerminalRole.Reset);
            int output = LogicChipGeometry.PinOf(unit, TerminalRole.Output);
            int discharge = LogicChipGeometry.PinOf(unit, TerminalRole.Discharge);
            bool trigger = LogicChipGeometry.IsLeft(chip, LogicChipGeometry.PinOf(unit, TerminalRole.Trigger));

            if (!trigger || !LogicChipGeometry.IsLeft(chip, output) || !LogicChipGeometry.IsLeft(chip, reset) ||
                LogicChipGeometry.IsLeft(chip, discharge) || LogicChipGeometry.IsLeft(chip, LogicChipGeometry.PinOf(unit, TerminalRole.Threshold)))
            {
                throw new InvalidOperationException($"{chip.Name}: a timer across the case has TRIG, OUT and RESET on the left, THRES, CONT and DISCH on the right.");
            }

            // The left side.
            var leftPart = new Drawing();
            LogicGateShape lower = LogicChipGeometry.LowerComparator(chip, unit, layout, edge + LogicChipGeometry.ComparatorBack, left: true);
            LogicGateShape stage = LogicChipGeometry.OutputStage(unit, layout, output, discharge: false);
            leftPart.Symbols.AddRange([lower, stage]);
            LogicChipGeometry.WireInputs([lower], edge, leftPart);
            leftPart.Wire(output, new Point(edge, Y(output)), stage.Output);

            // The right side, as if on the left.
            var rightPart = new Drawing();
            LogicGateShape upper = LogicChipGeometry.UpperComparator(unit, layout, edge + LogicChipGeometry.ComparatorBack);
            LogicGateShape drain = LogicChipGeometry.OutputStage(unit, layout, discharge, discharge: true);
            rightPart.Symbols.AddRange([upper, drain]);
            LogicChipGeometry.WireInputs([upper], edge, rightPart);
            rightPart.Wire(discharge, new Point(edge, Y(discharge)), drain.Output);

            // Inward of the right side's parts, where the flip-flop's right edge will be.
            double inward = edge + LogicChipGeometry.TimerFlipFlopAt;
            rightPart.Wire(0, upper.Output, new Point(inward, upper.CentreY));
            rightPart.Wire(0, new Point(inward, Y(discharge)), drain.Inputs[0]);

            double sum = body.Left + body.Right;
            drawing.Add(leftPart, null);
            drawing.Add(rightPart, sum);

            // The flip-flop between them.
            double top = Math.Min(lower.CentreY, Y(discharge)) - 12;
            double bottom = Math.Max(Y(reset), upper.CentreY) + 12;
            var rect = new Rect(inward, top, body.Width - (2 * LogicChipGeometry.TimerFlipFlopAt), bottom - top);

            LogicBlockPort s = LogicChipGeometry.PortOf(0, rect, BlockEdge.Left, lower.CentreY, "S");
            LogicBlockPort q = LogicChipGeometry.PortOf(0, rect, BlockEdge.Left, Y(output), "Q");
            LogicBlockPort clear = LogicChipGeometry.PortOf(reset, rect, BlockEdge.Left, Y(reset), "R", bubble: true);
            LogicBlockPort qBar = LogicChipGeometry.PortOf(0, rect, BlockEdge.Right, Y(discharge), "Q", overline: true);
            LogicBlockPort r = LogicChipGeometry.PortOf(0, rect, BlockEdge.Right, upper.CentreY, "R");

            drawing.Symbols.Add(new LogicBlockShape(rect, [unit], [s, q, clear, qBar, r]));
            drawing.Wire(0, lower.Output, s.At);
            drawing.Wire(0, q.At, stage.Inputs[0]);
            drawing.Wire(reset, new Point(edge, Y(reset)), clear.At);
        }
    }
}
