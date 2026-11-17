using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // The edge of a block a terminal is on.
    public enum BlockEdge
    {
        Left,
        Right,
        Top,
        Bottom,
    }

    // What a block shows in its middle: nothing, or a transceiver's two ways (the 245's).
    public enum BlockGlyph
    {
        None,
        Transceiver,
    }

    // ###########################################################################################
    // A terminal of a block: its pin (0 for one inside a unit), where its wire ends - outside its inverting bubble, when it
    // has one - the edge it is on, and what is written inside at it: its mark ("D", "Q", "EN", a weight "4"), the mark
    // overlined (Q-bar), a clock's wedge, and the mark of a 3-state or open-collector output.
    // ###########################################################################################
    public sealed record LogicBlockPort(
        int Pin,
        Point At,
        BlockEdge Edge,
        string Mark,
        bool Bubble = false,
        bool Clock = false,
        bool Overline = false,
        LogicOutputStage? Stage = null);

    // A block - a flip-flop, a latch, a decoder, a counter, a column's common control section, a transceiver: its rectangle,
    // the units it stands for, its terminals and what it shows in its middle.
    public sealed record LogicBlockShape(Rect Bounds, IReadOnlyList<LogicUnit> Units, IReadOnlyList<LogicBlockPort> Ports, BlockGlyph Glyph = BlockGlyph.None) : ILogicSymbol
    {
        public IReadOnlyList<LogicTerminalAt> Terminals => [.. this.Ports.Select(port => new LogicTerminalAt(port.Pin, port.At))];
    }

    // ###########################################################################################
    // THE CHIPS DRAWN AS BLOCKS (owner request, 2026-11-16: "OK, now go ahead and make the rest"), each a rectangle with its
    // terminals marked as the datasheets' logic symbols mark them - D and Q, Q-bar overlined, a clock's wedge, S and R, EN,
    // an active-low terminal's bubble, a decoder's address bits and outputs' numbers, a counter's bits - every wire straight
    // from its pin or bent once:
    //
    // - THE 7474'S FLIP-FLOPS as a flip-flop is drawn: D and the clock on its back, R on its top and S on its bottom from the
    //   pins above and below, Q and Q-bar from its front turning back to their pins under it, nested.
    // - FLIP-FLOPS AND LATCHES SHARING THEIR CONTROLS (74175, 74273, 74373) as the datasheets' logic symbols draw an array:
    //   a column per side, an element per flip-flop or latch with its D and Q at their pins, and a COMMON CONTROL SECTION at
    //   the top and bottom of each column with the shared clock, clear or enable - from its pin into the section on its
    //   side, and on across to the other column's.
    // - DECODERS AND THE COUNTER (74138, 74139, 74145, 74193): one block per unit - across the case when its pins are on
    //   both sides, else on its own side - each terminal on the edge facing its pin, at its pin.
    // ###########################################################################################
    public static partial class LogicChipGeometry
    {
        // From the case's edge to a block's outer side; a flip-flop's, a column's and a block's width.
        private const double BlockRun = 22;
        private const double FlipFlopWidth = 36;
        private const double ColumnWidth = 30;
        private const double ColumnGap = 40;
        private const double BlockWidth = 50;
        private const double SideBlockWidth = 34;

        // A terminal of `rect` on `edge`, `along` it (its y on a side, its x on the top or bottom).
        private static LogicBlockPort PortOf(
            int pin, Rect rect, BlockEdge edge, double along, string mark,
            bool bubble = false, bool clock = false, bool overline = false, LogicOutputStage? stage = null)
        {
            double ring = bubble ? LogicChipGeometry.Bubble : 0;

            Point at = edge switch
            {
                BlockEdge.Left => new Point(rect.Left - ring, along),
                BlockEdge.Right => new Point(rect.Right + ring, along),
                BlockEdge.Top => new Point(along, rect.Top - ring),
                _ => new Point(along, rect.Bottom + ring),
            };

            return new LogicBlockPort(pin, at, edge, mark, bubble, clock, overline, stage);
        }

        // ###########################################################################################
        // What is written at a block's terminal: a flip-flop's D and Q (Q-bar overlined), S and R, a clock's wedge or a
        // latch's C, EN, a decoder's address bits ("A0", "A1", ...) and outputs' numbers, a counter's "D0" to "D3" and "Q0" to
        // "Q3", its "+" and "-" clocks, LD, CO and BO - with the bubble of an active-low terminal and an output's 3-state or
        // open-collector mark. A weighted input is written by its bit's number, as many datasheets write them, not its weight:
        // a decoder's inputs "1", "2" beside its outputs "0" to "3" read as more outputs.
        // ###########################################################################################
        private static LogicBlockPort PortFor(LogicChip chip, LogicUnit unit, LogicTerminal terminal, Rect rect, BlockEdge edge, double along)
        {
            string index = terminal.Index.ToString(CultureInfo.InvariantCulture);
            string bit = terminal.Index > 0 ? System.Numerics.BitOperations.Log2((uint)terminal.Index).ToString(CultureInfo.InvariantCulture) : index;
            LogicOutputStage? stage = chip.Outputs == LogicOutputStage.TotemPole ? null : chip.Outputs;

            return terminal.Role switch
            {
                TerminalRole.Input when unit.Function == LogicFunction.Decoder => PortOf(terminal.Pin, rect, edge, along, "A" + bit),
                TerminalRole.Input when unit.Function == LogicFunction.Counter => PortOf(terminal.Pin, rect, edge, along, "D" + bit),
                TerminalRole.Input => PortOf(terminal.Pin, rect, edge, along, "D"),
                TerminalRole.Output when unit.Function == LogicFunction.Decoder => PortOf(terminal.Pin, rect, edge, along, index, bubble: terminal.ActiveLow, stage: stage),
                TerminalRole.Output when unit.Function == LogicFunction.Counter => PortOf(terminal.Pin, rect, edge, along, "Q" + bit),
                TerminalRole.Output => PortOf(terminal.Pin, rect, edge, along, "Q", overline: terminal.ActiveLow, stage: stage),
                TerminalRole.Clock when unit.Function == LogicFunction.Latch => PortOf(terminal.Pin, rect, edge, along, "C"),
                TerminalRole.Clock => PortOf(terminal.Pin, rect, edge, along, string.Empty, clock: true),
                TerminalRole.CountUp => PortOf(terminal.Pin, rect, edge, along, "+", clock: true),
                // Plain ASCII, as every user-visible mark is (code review, 2026-11-17: a Unicode minus until then).
                TerminalRole.CountDown => PortOf(terminal.Pin, rect, edge, along, "-", clock: true),
                TerminalRole.Set => PortOf(terminal.Pin, rect, edge, along, "S", bubble: terminal.ActiveLow),
                TerminalRole.Reset => PortOf(terminal.Pin, rect, edge, along, "R", bubble: terminal.ActiveLow),
                TerminalRole.Load => PortOf(terminal.Pin, rect, edge, along, "LD", bubble: terminal.ActiveLow),
                TerminalRole.Carry => PortOf(terminal.Pin, rect, edge, along, "CO", bubble: terminal.ActiveLow),
                TerminalRole.Borrow => PortOf(terminal.Pin, rect, edge, along, "BO", bubble: terminal.ActiveLow),
                TerminalRole.Enable => PortOf(terminal.Pin, rect, edge, along, "EN", bubble: terminal.ActiveLow),
                TerminalRole.Direction => PortOf(terminal.Pin, rect, edge, along, "DIR"),
                _ => PortOf(terminal.Pin, rect, edge, along, string.Empty, bubble: terminal.ActiveLow),
            };
        }

        // ---- the 7474's flip-flops ------------------------------------------------------------------

        private static LogicInside FlipFlopsFor(LogicChip chip, double nameWidth)
        {
            double reach = LogicChipGeometry.BlockRun + LogicChipGeometry.FlipFlopWidth + 18 + 6;
            PinoutLayout layout = LogicChipGeometry.Case(chip.PinCount, (2 * reach) + LogicChipGeometry.CentreGap, nameWidth);
            var drawing = new Drawing();

            foreach (LogicUnit unit in chip.Units)
            {
                bool left = LogicChipGeometry.IsLeft(chip, unit.Inputs[0]);
                var part = new Drawing();
                LogicChipGeometry.FlipFlopOnLeft(chip, unit, layout, part);
                drawing.Add(part, left ? null : layout.Body.Left + layout.Body.Right);
            }

            return drawing.In(layout);
        }

        // ###########################################################################################
        // A flip-flop whose pins are, down its side, its reset, D, clock, set, Q and Q-bar (the 7474's): D and the clock
        // straight into its back; the reset from the pin above onto its top edge and the set from the pin below onto its
        // bottom edge; Q-bar from its front's upper half and Q from its lower half, turning back to their pins under it -
        // the outer turn Q-bar's, so they nest. Worked out as if it stood on the left.
        // ###########################################################################################
        private static void FlipFlopOnLeft(LogicChip chip, LogicUnit unit, PinoutLayout layout, Drawing part)
        {
            double edge = layout.Body.Left;
            double Y(int pin) => LogicChipGeometry.YOf(layout, pin);

            LogicTerminal data = unit.Of(TerminalRole.Input).Single();
            LogicTerminal clock = unit.Of(TerminalRole.Clock).Single();
            LogicTerminal set = unit.Of(TerminalRole.Set).Single();
            LogicTerminal reset = unit.Of(TerminalRole.Reset).Single();
            LogicTerminal q = unit.Of(TerminalRole.Output).Single(terminal => !terminal.ActiveLow);
            LogicTerminal qBar = unit.Of(TerminalRole.Output).Single(terminal => terminal.ActiveLow);

            int[] rows = [.. new[] { reset, data, clock, set, q, qBar }.Select(terminal => LogicChipGeometry.RowOf(chip, terminal.Pin))];

            if (rows.Zip(rows.Skip(1), (above, below) => below - above).Any(step => step != 1))
            {
                throw new InvalidOperationException($"{chip.Name}: a flip-flop's pins must be its reset, D, clock, set, Q and Q-bar, one under the other.");
            }

            double x = edge + LogicChipGeometry.BlockRun;
            var rect = new Rect(x, Y(data.Pin) - 12, LogicChipGeometry.FlipFlopWidth, Y(clock.Pin) - Y(data.Pin) + 24);
            double middle = rect.Center.X;
            double inner = rect.Right + 9;
            double outer = rect.Right + 18;

            LogicBlockPort dataPort = LogicChipGeometry.PortFor(chip, unit, data, rect, BlockEdge.Left, Y(data.Pin));
            LogicBlockPort clockPort = LogicChipGeometry.PortFor(chip, unit, clock, rect, BlockEdge.Left, Y(clock.Pin));
            LogicBlockPort resetPort = LogicChipGeometry.PortFor(chip, unit, reset, rect, BlockEdge.Top, middle);
            LogicBlockPort setPort = LogicChipGeometry.PortFor(chip, unit, set, rect, BlockEdge.Bottom, middle);
            LogicBlockPort qBarPort = LogicChipGeometry.PortFor(chip, unit, qBar, rect, BlockEdge.Right, Y(data.Pin));
            LogicBlockPort qPort = LogicChipGeometry.PortFor(chip, unit, q, rect, BlockEdge.Right, Y(clock.Pin));

            part.Wire(data.Pin, new Point(edge, Y(data.Pin)), dataPort.At);
            part.Wire(clock.Pin, new Point(edge, Y(clock.Pin)), clockPort.At);
            part.Wire(reset.Pin, new Point(edge, Y(reset.Pin)), new Point(middle, Y(reset.Pin)), resetPort.At);
            part.Wire(set.Pin, new Point(edge, Y(set.Pin)), new Point(middle, Y(set.Pin)), setPort.At);
            part.Wire(qBar.Pin, qBarPort.At, new Point(outer, qBarPort.At.Y), new Point(outer, Y(qBar.Pin)), new Point(edge, Y(qBar.Pin)));
            part.Wire(q.Pin, qPort.At, new Point(inner, qPort.At.Y), new Point(inner, Y(q.Pin)), new Point(edge, Y(q.Pin)));

            part.Symbols.Add(new LogicBlockShape(rect, [unit], [dataPort, clockPort, resetPort, setPort, qBarPort, qPort]));
        }

        // ---- columns with their common control sections ----------------------------------------------------

        private static LogicInside ColumnsFor(LogicChip chip, double nameWidth)
        {
            PinoutLayout layout = LogicChipGeometry.Case(
                chip.PinCount, (2 * (LogicChipGeometry.BlockRun + LogicChipGeometry.ColumnWidth)) + LogicChipGeometry.ColumnGap, nameWidth);
            Rect body = layout.Body;
            double sum = body.Left + body.Right;

            // The controls every element shares - the clock, the clear, the enable - one of each.
            List<LogicTerminal> shared = chip.Units
                .SelectMany(unit => unit.Terminals)
                .GroupBy(terminal => terminal.Pin)
                .Where(group => group.Count() > 1)
                .Select(group => group.First())
                .ToList();

            var drawing = new Drawing();

            foreach (bool left in new[] { true, false })
            {
                List<LogicUnit> units = chip.Units.Where(unit => LogicChipGeometry.IsLeft(chip, unit.Inputs[0]) == left).ToList();
                var part = new Drawing();
                LogicChipGeometry.ColumnOnLeft(chip, units, shared, left, layout, part);
                drawing.Add(part, left ? null : sum);
            }

            // Each shared control on across, from the section on its pin's side to the other column's.
            List<LogicBlockPort> ports = drawing.Symbols.OfType<LogicBlockShape>().SelectMany(block => block.Ports).ToList();

            foreach (LogicTerminal control in shared)
            {
                LogicBlockPort leftEnd = ports.Single(port => port.Pin == control.Pin && port.Edge == BlockEdge.Right && port.At.X < body.Center.X);
                LogicBlockPort rightEnd = ports.Single(port => port.Pin == control.Pin && port.Edge == BlockEdge.Left && port.At.X > body.Center.X);
                drawing.Wire(control.Pin, leftEnd.At, rightEnd.At);
            }

            return drawing.In(layout);
        }

        // ###########################################################################################
        // One side's column, as if it stood on the left: an element per unit over the rows of its own pins, each terminal at
        // its pin, and the common control sections over the rows above and under them - a shared control whose pin is on
        // this side coming into its section here and leaving it by the inner edge, one from the other side coming in by the
        // inner edge, marked there as it is at its pin.
        // ###########################################################################################
        private static void ColumnOnLeft(LogicChip chip, IReadOnlyList<LogicUnit> units, IReadOnlyList<LogicTerminal> shared, bool left, PinoutLayout layout, Drawing part)
        {
            Rect body = layout.Body;
            double edge = body.Left;
            double x = edge + LogicChipGeometry.BlockRun;
            HashSet<int> sharedPins = shared.Select(control => control.Pin).ToHashSet();
            double Y(int pin) => LogicChipGeometry.YOf(layout, pin);

            int first = int.MaxValue;
            int last = int.MinValue;

            foreach (LogicUnit unit in units)
            {
                List<LogicTerminal> own = unit.Terminals.Where(terminal => !sharedPins.Contains(terminal.Pin)).ToList();
                int top = own.Min(terminal => LogicChipGeometry.RowOf(chip, terminal.Pin));
                int bottom = own.Max(terminal => LogicChipGeometry.RowOf(chip, terminal.Pin));
                first = Math.Min(first, top);
                last = Math.Max(last, bottom);

                double rectTop = LogicChipGeometry.RowY(layout, top - 0.5);
                var rect = new Rect(x, rectTop, LogicChipGeometry.ColumnWidth, LogicChipGeometry.RowY(layout, bottom + 0.5) - rectTop);
                var ports = new List<LogicBlockPort>();

                foreach (LogicTerminal terminal in own)
                {
                    LogicBlockPort port = LogicChipGeometry.PortFor(chip, unit, terminal, rect, BlockEdge.Left, Y(terminal.Pin));
                    ports.Add(port);
                    part.Wire(terminal.Pin, new Point(edge, Y(terminal.Pin)), port.At);
                }

                part.Symbols.Add(new LogicBlockShape(rect, [unit], ports));
            }

            double firstTop = LogicChipGeometry.RowY(layout, first - 0.5);
            double lastBottom = LogicChipGeometry.RowY(layout, last + 0.5);

            foreach (bool above in new[] { true, false })
            {
                List<LogicTerminal> controls = shared.Where(control => (LogicChipGeometry.RowOf(chip, control.Pin) < first) == above).ToList();

                if (controls.Count == 0)
                {
                    continue;
                }

                // Over the rows of its controls, as an element is over its own - so it stands clear of the case's end as the
                // pins' rows do (owner request, 2026-11-16: "a little vertical/empty spacing between the top/bottom gate
                // borders and the frame of the IC").
                List<int> rows = controls.Select(control => LogicChipGeometry.RowOf(chip, control.Pin)).ToList();

                Rect rect = above
                    ? new Rect(x, LogicChipGeometry.RowY(layout, rows.Min() - 0.5), LogicChipGeometry.ColumnWidth, firstTop - LogicChipGeometry.RowY(layout, rows.Min() - 0.5))
                    : new Rect(x, lastBottom, LogicChipGeometry.ColumnWidth, LogicChipGeometry.RowY(layout, rows.Max() + 0.5) - lastBottom);
                var ports = new List<LogicBlockPort>();

                foreach (LogicTerminal control in controls)
                {
                    double y = Y(control.Pin);

                    if (LogicChipGeometry.IsLeft(chip, control.Pin) == left)
                    {
                        LogicBlockPort port = LogicChipGeometry.PortFor(chip, units[0], control, rect, BlockEdge.Left, y);
                        ports.Add(port);
                        ports.Add(LogicChipGeometry.PortOf(control.Pin, rect, BlockEdge.Right, y, string.Empty));
                        part.Wire(control.Pin, new Point(edge, y), port.At);
                    }
                    else
                    {
                        ports.Add(LogicChipGeometry.PortFor(chip, units[0], control, rect, BlockEdge.Right, y));
                    }
                }

                part.Symbols.Add(new LogicBlockShape(rect, units, ports));
            }
        }

        // ---- decoders and the counter -------------------------------------------------------------------

        private static LogicInside BlocksFor(LogicChip chip, double nameWidth)
        {
            bool across = chip.Units.Any(unit => unit.Pins.Select(pin => LogicChipGeometry.IsLeft(chip, pin)).Distinct().Count() > 1);

            if (across && chip.Units.Count > 1)
            {
                throw new InvalidOperationException($"{chip.Name}: only a chip of one unit is drawn as a block across the case.");
            }

            double width = across
                ? (2 * LogicChipGeometry.BlockRun) + LogicChipGeometry.BlockWidth
                : (2 * (LogicChipGeometry.BlockRun + LogicChipGeometry.SideBlockWidth)) + LogicChipGeometry.CentreGap;
            PinoutLayout layout = LogicChipGeometry.Case(chip.PinCount, width, nameWidth);
            Rect body = layout.Body;
            var drawing = new Drawing();

            foreach (LogicUnit unit in chip.Units)
            {
                int top = unit.Pins.Min(pin => LogicChipGeometry.RowOf(chip, pin));
                int bottom = unit.Pins.Max(pin => LogicChipGeometry.RowOf(chip, pin));
                double rectTop = LogicChipGeometry.RowY(layout, top - 0.5) + 2;
                double rectHeight = LogicChipGeometry.RowY(layout, bottom + 0.5) - 2 - rectTop;
                bool left = LogicChipGeometry.IsLeft(chip, unit.Pins.First());

                Rect rect = across
                    ? new Rect(body.Left + LogicChipGeometry.BlockRun, rectTop, LogicChipGeometry.BlockWidth, rectHeight)
                    : left
                        ? new Rect(body.Left + LogicChipGeometry.BlockRun, rectTop, LogicChipGeometry.SideBlockWidth, rectHeight)
                        : new Rect(body.Right - LogicChipGeometry.BlockRun - LogicChipGeometry.SideBlockWidth, rectTop, LogicChipGeometry.SideBlockWidth, rectHeight);

                var ports = new List<LogicBlockPort>();

                foreach (LogicTerminal terminal in unit.Terminals)
                {
                    bool pinLeft = LogicChipGeometry.IsLeft(chip, terminal.Pin);

                    if (!across && pinLeft != left)
                    {
                        throw new InvalidOperationException($"{chip.Name}: a block on one side cannot reach pin {terminal.Pin} on the other.");
                    }

                    double y = LogicChipGeometry.YOf(layout, terminal.Pin);
                    LogicBlockPort port = LogicChipGeometry.PortFor(chip, unit, terminal, rect, pinLeft ? BlockEdge.Left : BlockEdge.Right, y);
                    ports.Add(port);
                    drawing.Wire(terminal.Pin, new Point(pinLeft ? body.Left : body.Right, y), port.At);
                }

                drawing.Symbols.Add(new LogicBlockShape(rect, [unit], ports));
            }

            return drawing.In(layout);
        }
    }
}
