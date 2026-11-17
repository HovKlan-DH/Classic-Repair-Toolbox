using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // Which way a gate points: its output to the right (it stands on the left side's pins) or to the left.
    public enum GateFacing
    {
        Right,
        Left,
    }

    // The shapes a gate is drawn in - a comparator is a triangle with "+" and "-" at its inputs (a timer's).
    public enum GateSymbol
    {
        And,
        Or,
        Xor,
        Buffer,
        Mux,
        Comparator,
    }

    // Where a wire meets a symbol, and the pin whose net it is - 0 for a connection inside a unit (a timer's comparator to
    // its flip-flop), which reaches no pin.
    public sealed record LogicTerminalAt(int Pin, Point At);

    // ###########################################################################################
    // Anything drawn inside the case that wires meet: a gate, a block, a switch. Bounds is what no wire may run through;
    // Units the chip's units it stands for, whole or in part (a column's common control section stands for every latch of
    // the column); Terminals where wires meet it. Two terminals of one symbol with the same pin are the same net, passed
    // through it (a section's control, on to the other column).
    // ###########################################################################################
    public interface ILogicSymbol
    {
        Rect Bounds { get; }

        IReadOnlyList<LogicUnit> Units { get; }

        IReadOnlyList<LogicTerminalAt> Terminals { get; }
    }

    // A select or enable on a gate's top or bottom edge: where its wire ends, where it meets the edge, and the inverting
    // bubble between the two (an active-low enable's).
    public sealed record LogicGateControl(Point Terminal, Point Edge, bool Bubble);

    // ###########################################################################################
    // One gate, multiplexer, buffer or comparator as drawn. Bounds is the whole symbol, its bubbles included; BackX the x of
    // its back, where its inputs' wires end; BodyHeight the shape's own height about CentreY, BackTop and BackBottom how far
    // its back line reaches (further than the shape for a gate with many inputs - the datasheets' extended back line).
    // Inputs are where each input's wire ends, with the mark written at it (a multiplexer's "0" to "3", a comparator's "+",
    // "-" and its internal "1/3"); Output where the output's wire starts; Controls its selects and enable on its top or
    // bottom edge. How it draws - inverted, its output stage, Schmitt-trigger inputs - is its own, as a timer's parts differ.
    // ###########################################################################################
    public sealed record LogicGateShape(
        GateSymbol Shape,
        bool Inverted,
        LogicOutputStage Stage,
        bool Schmitt,
        GateFacing Facing,
        Rect Bounds,
        double BackX,
        double BodyWidth,
        double BodyHeight,
        double CentreY,
        double BackTop,
        double BackBottom,
        IReadOnlyList<Point> Inputs,
        IReadOnlyList<string> InputMarks,
        Point Output,
        IReadOnlyList<LogicGateControl> Controls,
        IReadOnlyList<LogicUnit> Units,
        IReadOnlyList<LogicTerminalAt> Terminals) : ILogicSymbol
    {
        // The unit it is - a gate stands for one.
        public LogicUnit Unit => this.Units[0];

        // What is written outside its back at an input with no wire, one per input ("" for none) - a timer's comparator's
        // internal "1/3" of the supply.
        public IReadOnlyList<string> InputTags { get; init; } = [];
    }

    // A wire inside the case, of `Pin`'s net - 0 for one inside a unit - as the points it runs through.
    public sealed record LogicWire(int Pin, IReadOnlyList<Point> Points);

    // ###########################################################################################
    // A pin's number on a badge at a symbol, in the pin's color as its square is: what a stub from that symbol is joined to -
    // an input from the other side of the case, a shared select or enable. Stub is where the stub's wire meets the badge,
    // on the middle of its side.
    // ###########################################################################################
    public sealed record LogicPinBadge(int Pin, Rect Area, Point Stub);

    // The case of a chip drawn with its insides: the case and its pins (PinoutGeometry's record, so the pointer finds the
    // pins as it does on any case), the symbols, the wires and the badges.
    public sealed record LogicInside(
        PinoutLayout Case,
        IReadOnlyList<ILogicSymbol> Symbols,
        IReadOnlyList<LogicWire> Wires,
        IReadOnlyList<LogicPinBadge> Badges);

    // ###########################################################################################
    // THE INSIDES OF A CHIP, DRAWN IN ITS CASE (owner request, 2026-11-16: "visualize its internals for e.g. its logic
    // cirquits - something alike the attached image, just better and prettier"; LogicChips decides which chip, PinoutDrawing
    // paints). The case is a DIP seen from above as PinoutGeometry draws it - pins 1 to N/2 down the left, the rest up the
    // right - but wider and taller, each pin a square outside it with its number, so the insides fit between the pins.
    //
    // EVERY WIRE IS STRAIGHT OR BENDS AT RIGHT ANGLES, AND NO TWO CROSS. What does not fit so is not wired to its pin: a
    // stub ends on a BADGE with the pin's number, in its color (LogicPinBadge) - an input from the other side of the case,
    // a control several units share. Each chip family is laid out its own way (LogicChip.Layout), every one of them by these
    // rules - the gates here, the others in their parts:
    //
    //   LogicChipGeometry.cs            this - the records, the case, the gates and multiplexers (and their selects and
    //                                   enables), the helpers the other parts share
    //   LogicChipGeometry.Blocks.cs     flip-flops, latches and flip-flops in columns, decoders and the counter
    //   LogicChipGeometry.Arrays.cs     the octal buffers and the bus transceiver, across the case
    //   LogicChipGeometry.Switches.cs   the 4066's switches
    //   LogicChipGeometry.Timers.cs     the 555's and 556's comparators, flip-flop and output stages
    //   LogicChipGeometry.Marks.cs      the marks inside a gate - signs, numbers, stage and Schmitt marks - kept clear of
    //                                   its outline, and how large a gate is for them
    //
    // NOTHING IS WRITTEN AT THE PINS INSIDE THE CASE (owner request, the same day: "I do not think this text inside the IC is
    // needed (e.g. "GND" and "VCC") as there will always be a text for the pin itself") - the pin's name stands outside it
    // (the datasheet's where the board names none - PinoutDrawing), and a pin no unit uses, the supply and one not
    // connected, has no wire inside at all. What a block's terminal is - "D", "Q", "EN", a clock's wedge - is its symbol's.
    //
    // LogicChipGeometryTests read the connections back off the drawing - wire ends on pins, terminals and badges - and hold
    // them against the chip's units, no wire or badge touching another or running through a symbol.
    //
    // Coordinates start at (0, 0), as PinoutGeometry's; every size here is for a pitch of 24 between pins.
    // ###########################################################################################
    public static partial class LogicChipGeometry
    {
        public const double Pitch = 24;
        public const double EndPadding = 10;
        public const double Square = 18;
        public const double LeadGap = 5;

        // ###########################################################################################
        // ROOM BETWEEN THE INSIDES AND THE CASE (owner request, 2026-11-16: "a little vertical/empty spacing between the
        // top/bottom gate borders and the frame of the IC ... I want to stand out for these pinout IC's"): no symbol comes
        // nearer any edge of the case than FrameClear, no badge nearer than BadgeFrameClear, and neither comes near the notch
        // (NotchRadius, round the middle of the top edge) or the dot by pin 1 (FirstPinDot from the top left corner, DotRadius)
        // - which PinoutDrawing paints by these same numbers. LogicChipGeometryTests hold every chip to it.
        // ###########################################################################################
        public const double FrameClear = 8;
        public const double BadgeFrameClear = 6;
        public const double NotchRadius = 7;
        public const double FirstPinDot = 9;
        public const double DotRadius = 3;

        // From the case's edge to a gate's back - room for an input's badge beyond it - and from its output to the bend back
        // to its pin, or across.
        public const double InputRun = 27;
        public const double OutputRun = 10;
        public const double JogRun = 16;

        // Between the furthest reaches of the two sides' gates, and the least a side reaches with no gate on it.
        public const double CentreGap = 14;
        public const double LeastReach = 30;

        public const double Bubble = 6;

        // How thick a gate's outline is drawn (PinoutDrawing) - what a mark inside it keeps clear of.
        public const double GateOutline = 1.5;

        // A pin's badge at a symbol, and the stub to it from an input's end on a gate's back; how far clear of a symbol a
        // badge stands.
        public const double BadgeWidth = 15;
        public const double BadgeHeight = 13;
        public const double InputStub = 4;
        private const double BadgeClear = 3;

        // A multiplexer's selects along its top edge from its back, one apart from the next, and its enable that far from its
        // front; their stubs up from it, before they turn - a select's to its badge beyond the back, the enable's to its
        // badge beyond the front.
        private const double SelectAt = 9;
        private const double SelectApart = 8;
        private const double EnableFromFront = 10;
        private const double BufferEnableAt = 5;
        private const double SelectStub = 6;
        private const double EnableStub = 4;

        // How much lower a multiplexer's top edge is at its front than at its back (and its bottom edge higher): a fifth of its
        // height, but no more than 8 - the 153's tall one would otherwise slope so steeply that its selects meet it deep
        // inside its outline.
        public static double MuxTaper(double height) => Math.Min(0.2 * height, 8);

        // A shape's width, without its bubble.
        public static double BodyWidthOf(GateSymbol shape) => shape switch
        {
            GateSymbol.Buffer => 30,
            GateSymbol.Mux => 32,
            GateSymbol.Comparator => 30,
            GateSymbol.Xor => 39,
            _ => 34,
        };

        // ###########################################################################################
        // The case of `chip` with its insides, its name columns `nameWidth` wide on each side (PinoutGeometry.NameWidthFor).
        // ###########################################################################################
        public static LogicInside For(LogicChip chip, double nameWidth = PinoutGeometry.NameWidth)
        {
            ArgumentNullException.ThrowIfNull(chip);

            return chip.Layout switch
            {
                LogicLayout.Buffers => LogicChipGeometry.BuffersFor(chip, nameWidth),
                LogicLayout.Transceivers => LogicChipGeometry.TransceiversFor(chip, nameWidth),
                LogicLayout.Switches => LogicChipGeometry.SwitchesFor(chip, nameWidth),
                LogicLayout.FlipFlops => LogicChipGeometry.FlipFlopsFor(chip, nameWidth),
                LogicLayout.Columns => LogicChipGeometry.ColumnsFor(chip, nameWidth),
                LogicLayout.Blocks => LogicChipGeometry.BlocksFor(chip, nameWidth),
                LogicLayout.Timers => LogicChipGeometry.TimersFor(chip, nameWidth),
                _ => LogicChipGeometry.GatesFor(chip, nameWidth),
            };
        }

        // ###########################################################################################
        // *** EVERY DIP IS DRAWN IN THIS CASE (owner request, 2026-11-17: "unify ALL pinout IC's so they look the same,
        // meaning that IC's where it has no data (like e.g. the CPU or VIC-II etc) will look like the new ones you have
        // generated, so we can get the same directions applied to it"). *** One CRT knows nothing inside of - a CPU, a ROM, a
        // custom chip - is this case with nothing in it: each pin a square in its direction's color with its number, its name
        // outside, the notch and pin 1's dot - its body as wide as a real one's in proportion to its pins, 100 mil apart: 300 mil,
        // three pitches, or 600 mil, six - the latter for a package named "wide" or of 24 pins or more (the SID, the ROMs,
        // every 40-pin chip). Null for a package that is no DIP CRT can count: a single row is PinoutGeometry's.
        // ###########################################################################################
        public static LogicInside? EmptyCaseFor(string? package, double nameWidth = PinoutGeometry.NameWidth)
        {
            string text = (package ?? string.Empty).Trim();

            if (!ComponentVocabulary.IsDualInLine(text) || ComponentVocabulary.PinCountOf(text) is not int count || count < 2 || count % 2 != 0)
            {
                return null;
            }

            bool wide = text.Contains("wide", StringComparison.OrdinalIgnoreCase) || count >= LogicChipGeometry.WideFrom;

            return new LogicInside(LogicChipGeometry.Case(count, (wide ? 6 : 3) * LogicChipGeometry.Pitch, nameWidth), [], [], []);
        }

        // The pin count from which an empty case is a 600 mil one.
        public const int WideFrom = 24;

        // ###########################################################################################
        // A DIP of `count` pins with its body `bodyWidth` wide and `endPadding` above its first pin's row and under its last:
        // each pin a square beside the body, its number in it, its name outside it in a column `nameWidth` wide.
        // ###########################################################################################
        public static PinoutLayout Case(int count, double bodyWidth, double nameWidth = PinoutGeometry.NameWidth, double endPadding = LogicChipGeometry.EndPadding)
        {
            if (count < 2 || count % 2 != 0)
            {
                throw new ArgumentOutOfRangeException(nameof(count), count, "A DIP has an even number of pins, at least two.");
            }

            int perSide = count / 2;
            double bodyLeft = nameWidth + LogicChipGeometry.Square + LogicChipGeometry.LeadGap;
            var body = new Rect(bodyLeft, 0, bodyWidth, (perSide * LogicChipGeometry.Pitch) + (2 * endPadding));
            var pins = new List<PinoutPin>(count);

            for (int index = 0; index < count; index++)
            {
                bool left = index < perSide;
                int row = left ? index : count - 1 - index;
                double centreY = body.Top + endPadding + ((row + 0.5) * LogicChipGeometry.Pitch);
                double half = LogicChipGeometry.Square / 2;

                var square = left
                    ? new Rect(body.Left - LogicChipGeometry.LeadGap - LogicChipGeometry.Square, centreY - half, LogicChipGeometry.Square, LogicChipGeometry.Square)
                    : new Rect(body.Right + LogicChipGeometry.LeadGap, centreY - half, LogicChipGeometry.Square, LogicChipGeometry.Square);

                pins.Add(new PinoutPin(
                    index + 1,
                    left ? PinSide.Left : PinSide.Right,
                    square,
                    square.Center,
                    left ? new Point(square.Left - 4, centreY) : new Point(square.Right + 4, centreY)));
            }

            var size = new Size(body.Right + LogicChipGeometry.LeadGap + LogicChipGeometry.Square + nameWidth, body.Bottom);

            return new PinoutLayout(body, pins, size, LogicChipGeometry.Pitch);
        }

        // ---- shared helpers ------------------------------------------------------------------------

        private static bool IsLeft(LogicChip chip, int pin) => pin <= chip.PinCount / 2;

        // A pin's row on its side, 0 the top.
        private static int RowOf(LogicChip chip, int pin) => LogicChipGeometry.IsLeft(chip, pin) ? pin - 1 : chip.PinCount - pin;

        // The pin at `row` of a side.
        private static int PinAtRow(LogicChip chip, int row, bool left) => left ? row + 1 : chip.PinCount - row;

        // The y of a row - between two pins for a half, beyond the first or last for one before or after them.
        private static double RowY(PinoutLayout layout, double row) =>
            layout.Pins[0].Lead.Center.Y + (row * LogicChipGeometry.Pitch);

        private static double YOf(PinoutLayout layout, int pin) => layout.Pins[pin - 1].Lead.Center.Y;

        // Every pin some unit uses.
        private static HashSet<int> UsedPins(LogicChip chip) => chip.Units.SelectMany(unit => unit.Pins).ToHashSet();

        // A badge whose right side's middle is `stub`, and one whose left side's middle is.
        private static Rect BadgeEndingAt(Point stub) =>
            new(stub.X - LogicChipGeometry.BadgeWidth, stub.Y - (LogicChipGeometry.BadgeHeight / 2), LogicChipGeometry.BadgeWidth, LogicChipGeometry.BadgeHeight);

        private static Rect BadgeStartingAt(Point stub) =>
            new(stub.X, stub.Y - (LogicChipGeometry.BadgeHeight / 2), LogicChipGeometry.BadgeWidth, LogicChipGeometry.BadgeHeight);

        // ###########################################################################################
        // What a layout puts in the case, built side by side: a unit on the right is worked out as if it stood on the left -
        // pointing right, its pins' edge the left - and mirrored about the case's middle as it is added.
        // ###########################################################################################
        private sealed class Drawing
        {
            public List<ILogicSymbol> Symbols { get; } = [];

            public List<LogicWire> Wires { get; } = [];

            public List<LogicPinBadge> Badges { get; } = [];

            public void Wire(int pin, params Point[] points) => this.Wires.Add(new LogicWire(pin, points));

            // `part`, as it is or mirrored about the line x = sum / 2.
            public void Add(Drawing part, double? mirrorSum)
            {
                if (mirrorSum is not double sum)
                {
                    this.Symbols.AddRange(part.Symbols);
                    this.Wires.AddRange(part.Wires);
                    this.Badges.AddRange(part.Badges);
                    return;
                }

                this.Symbols.AddRange(part.Symbols.Select(symbol => LogicChipGeometry.Mirrored(symbol, sum)));
                this.Wires.AddRange(part.Wires.Select(wire => wire with { Points = [.. wire.Points.Select(point => Mirrored(point, sum))] }));
                this.Badges.AddRange(part.Badges.Select(badge => badge with { Area = Mirrored(badge.Area, sum), Stub = Mirrored(badge.Stub, sum) }));
            }

            public LogicInside In(PinoutLayout layout) => new(layout, this.Symbols, this.Wires, this.Badges);
        }

        private static Point Mirrored(Point point, double sum) => new(sum - point.X, point.Y);

        private static Rect Mirrored(Rect rect, double sum) => new(sum - rect.Right, rect.Top, rect.Width, rect.Height);

        private static ILogicSymbol Mirrored(ILogicSymbol symbol, double sum) => symbol switch
        {
            LogicGateShape gate => gate with
            {
                Facing = gate.Facing == GateFacing.Right ? GateFacing.Left : GateFacing.Right,
                Bounds = Mirrored(gate.Bounds, sum),
                BackX = sum - gate.BackX,
                Inputs = [.. gate.Inputs.Select(point => Mirrored(point, sum))],
                Output = Mirrored(gate.Output, sum),
                Controls = [.. gate.Controls.Select(control => control with { Terminal = Mirrored(control.Terminal, sum), Edge = Mirrored(control.Edge, sum) })],
                Terminals = [.. gate.Terminals.Select(terminal => terminal with { At = Mirrored(terminal.At, sum) })],
            },
            LogicBlockShape block => block with
            {
                Bounds = Mirrored(block.Bounds, sum),
                Ports = [.. block.Ports.Select(port => port with
                {
                    At = Mirrored(port.At, sum),
                    Edge = port.Edge switch { BlockEdge.Left => BlockEdge.Right, BlockEdge.Right => BlockEdge.Left, _ => port.Edge },
                })],
            },
            LogicSwitchShape blade => blade with
            {
                Pivot = Mirrored(blade.Pivot, sum),
                Contact = Mirrored(blade.Contact, sum),
                BladeTip = Mirrored(blade.BladeTip, sum),
                Control = Mirrored(blade.Control, sum),
                Link = Mirrored(blade.Link, sum),
                Bounds = Mirrored(blade.Bounds, sum),
                Terminals = [.. blade.Terminals.Select(terminal => terminal with { At = Mirrored(terminal.At, sum) })],
            },
            _ => throw new ArgumentOutOfRangeException(nameof(symbol), symbol, "A symbol LogicChipGeometry does not know."),
        };

        // ---- the gates ---------------------------------------------------------------------------

        // The shape a gate unit is drawn in.
        private static GateSymbol ShapeOf(LogicFunction function) => function switch
        {
            LogicFunction.And => GateSymbol.And,
            LogicFunction.Or => GateSymbol.Or,
            LogicFunction.Xor => GateSymbol.Xor,
            LogicFunction.Buffer => GateSymbol.Buffer,
            _ => GateSymbol.Mux,
        };

        private static double GateWidthOf(LogicUnit unit) =>
            unit.Function == LogicFunction.Selector ? 36 : LogicChipGeometry.BodyWidthOf(LogicChipGeometry.ShapeOf(unit.Function));

        // ###########################################################################################
        // EACH GATE STANDS ON THE SIDE MOST OF ITS INPUTS ARE ON, pointing into the case: its inputs run straight in from their
        // pins, and its output turns back to its pin beside them (as every datasheet's connection diagram draws a 7400) - or,
        // on the other side, runs across to it. An input from the other side (the 7410's 1C, the 74133's H to M) ends on its
        // pin's badge. A multiplexer's or buffer's select and enable come in on its top edge - from their pin when it is just
        // above the gate on its side (the 125's enable, the 153's), else from a badge.
        // ###########################################################################################
        private static LogicInside GatesFor(LogicChip chip, double nameWidth)
        {
            List<(LogicUnit Unit, bool Left)> homes = chip.Units.Select(unit => (unit, LogicChipGeometry.HomeIsLeft(chip, unit))).ToList();

            double leftReach = LogicChipGeometry.ReachOf(chip, homes, left: true);
            double rightReach = LogicChipGeometry.ReachOf(chip, homes, left: false);
            PinoutLayout layout = LogicChipGeometry.Case(chip.PinCount, leftReach + rightReach + LogicChipGeometry.CentreGap, nameWidth);
            var drawing = new Drawing();

            foreach ((LogicUnit unit, bool left) in homes)
            {
                var part = new Drawing();
                LogicChipGeometry.PlaceGate(chip, unit, left, layout, part);
                drawing.Add(part, left ? null : layout.Body.Left + layout.Body.Right);
            }

            return drawing.In(layout);
        }

        // The side most of a unit's inputs are on; a tie goes to its output's side.
        private static bool HomeIsLeft(LogicChip chip, LogicUnit unit)
        {
            IReadOnlyList<int> inputs = unit.Inputs;
            int onLeft = inputs.Count(pin => LogicChipGeometry.IsLeft(chip, pin));
            int onRight = inputs.Count - onLeft;

            return onLeft != onRight ? onLeft > onRight : LogicChipGeometry.IsLeft(chip, unit.Output);
        }

        // ###########################################################################################
        // How far into the case a side's gates reach from its edge - past a gate's front its output's bend, or across, and an
        // enable's badge beyond it - and at least LeastReach, so a side with none has some case.
        // ###########################################################################################
        private static double ReachOf(LogicChip chip, IReadOnlyList<(LogicUnit Unit, bool Left)> homes, bool left) =>
            homes
                .Where(home => home.Left == left)
                .Select(home =>
                    LogicChipGeometry.InputRun +
                    LogicChipGeometry.GateWidthOf(home.Unit) +
                    (home.Unit.Inverted ? LogicChipGeometry.Bubble : 0) +
                    Math.Max(
                        LogicChipGeometry.IsLeft(chip, home.Unit.Output) == left ? LogicChipGeometry.OutputRun : LogicChipGeometry.JogRun,
                        LogicChipGeometry.ControlsOf(chip, home.Unit, left).Any(control => control.Way == ControlWay.Badge && control.Terminal.Role == TerminalRole.Enable)
                            ? BadgeClear + LogicChipGeometry.BadgeWidth
                            : 0))
                .Append(LogicChipGeometry.LeastReach)
                .Max();

        // How a select or enable reaches its gate's edge: from its pin above the gate, from its pin below it, or by its badge.
        private enum ControlWay
        {
            Above,
            Below,
            Badge,
        }

        // ###########################################################################################
        // A gate's selects (by weight) and its enable, each with the way it comes in and where it meets the edge (from the
        // back): FROM ITS PIN when the pin is on the gate's side above (or below) its inputs, with every row between free or
        // another of its own controls, and the output not turning back past it - else by its badge, a select's beyond the
        // back, an enable's beyond the front. ALONG THE TOP EDGE their order keeps every wire clear of the others: first the
        // selects going back to their badges, then those from their pins - the nearest pin first, so the further a pin, the
        // further toward the front its wire comes down, and they nest - then the enable going on to its badge, at its own
        // place near the front. Below, those from their pins alike.
        // ###########################################################################################
        private static List<(LogicTerminal Terminal, ControlWay Way, double X)> ControlsOf(LogicChip chip, LogicUnit unit, bool left)
        {
            var controls = unit.Of(TerminalRole.Select).OrderBy(terminal => terminal.Index).Concat(unit.Of(TerminalRole.Enable)).ToList();

            if (controls.Count == 0)
            {
                return [];
            }

            GateSymbol shape = LogicChipGeometry.ShapeOf(unit.Function);
            double width = LogicChipGeometry.GateWidthOf(unit);
            List<int> inputRows = unit.Inputs.Where(pin => LogicChipGeometry.IsLeft(chip, pin) == left).Select(pin => LogicChipGeometry.RowOf(chip, pin)).ToList();
            int top = inputRows.Min();
            int bottom = inputRows.Max();
            HashSet<int> used = LogicChipGeometry.UsedPins(chip);
            HashSet<int> own = controls.Select(terminal => terminal.Pin).ToHashSet();
            bool outputHere = LogicChipGeometry.IsLeft(chip, unit.Output) == left;
            int outputRow = LogicChipGeometry.RowOf(chip, unit.Output);

            bool RowsFree(int from, int to) =>
                Enumerable.Range(Math.Min(from, to) + 1, Math.Max(0, Math.Abs(to - from) - 1))
                    .Select(row => LogicChipGeometry.PinAtRow(chip, row, left))
                    .All(pin => !used.Contains(pin) || own.Contains(pin));

            ControlWay WayOf(LogicTerminal control)
            {
                if (LogicChipGeometry.IsLeft(chip, control.Pin) != left)
                {
                    return ControlWay.Badge;
                }

                int row = LogicChipGeometry.RowOf(chip, control.Pin);

                return row < top && RowsFree(row, top) && !(outputHere && outputRow < row) ? ControlWay.Above
                    : row > bottom && RowsFree(bottom, row) && !(outputHere && outputRow > row) ? ControlWay.Below
                    : ControlWay.Badge;
            }

            List<(LogicTerminal Terminal, ControlWay Way)> ways = controls.Select(control => (control, WayOf(control))).ToList();
            var placed = new List<(LogicTerminal Terminal, ControlWay Way, double X)>();
            double start = shape == GateSymbol.Buffer ? LogicChipGeometry.BufferEnableAt : LogicChipGeometry.SelectAt;

            // Along the top edge: the selects to the back's badges, then from above - nearest first.
            double x = start;

            foreach ((LogicTerminal control, ControlWay way) in ways.Where(item => item.Way == ControlWay.Badge && item.Terminal.Role == TerminalRole.Select)
                         .Concat(ways.Where(item => item.Way == ControlWay.Above).OrderByDescending(item => LogicChipGeometry.RowOf(chip, item.Terminal.Pin))))
            {
                placed.Add((control, way, x));
                x += LogicChipGeometry.SelectApart;
            }

            // The enable to its badge beyond the front.
            foreach ((LogicTerminal control, ControlWay way) in ways.Where(item => item.Way == ControlWay.Badge && item.Terminal.Role != TerminalRole.Select))
            {
                placed.Add((control, way, Math.Max(x, width - LogicChipGeometry.EnableFromFront)));
            }

            // Along the bottom edge: from below - nearest first.
            x = start;

            foreach ((LogicTerminal control, ControlWay way) in ways.Where(item => item.Way == ControlWay.Below).OrderBy(item => LogicChipGeometry.RowOf(chip, item.Terminal.Pin)))
            {
                placed.Add((control, way, x));
                x += LogicChipGeometry.SelectApart;
            }

            return placed;
        }

        // ###########################################################################################
        // One unit, worked out as if it stood on the left (pointing right) and mirrored for the right: its inputs where
        // their pins are - an input from the other side halfway between two of them, on a stub to its pin's badge - its
        // height from how far apart they are, its output turned back to its pin, or across to the other side, and its
        // selects and enable on its top (or bottom) edge.
        // ###########################################################################################
        private static void PlaceGate(LogicChip chip, LogicUnit unit, bool left, PinoutLayout layout, Drawing part)
        {
            Rect body = layout.Body;
            double edge = body.Left;
            double farEdge = body.Right;
            double Y(int pin) => LogicChipGeometry.YOf(layout, pin);

            List<LogicTerminal> inputs = unit.Of(TerminalRole.Input).ToList();

            // Each input's height: its pin's for one on this side, else the next free halfway point between them.
            List<double> wired = inputs.Where(terminal => LogicChipGeometry.IsLeft(chip, terminal.Pin) == left).Select(terminal => Y(terminal.Pin)).OrderBy(y => y).ToList();
            var gaps = new Queue<double>(wired.Zip(wired.Skip(1), (above, below) => (above + below) / 2));
            var slots = new List<(LogicTerminal Terminal, double Y, bool Wired)>();

            foreach (LogicTerminal terminal in inputs)
            {
                if (LogicChipGeometry.IsLeft(chip, terminal.Pin) == left)
                {
                    slots.Add((terminal, Y(terminal.Pin), true));
                }
                else if (gaps.Count > 0)
                {
                    slots.Add((terminal, gaps.Dequeue(), false));
                }
                else
                {
                    throw new InvalidOperationException($"{chip.Name}: no room on its gate for input {terminal.Pin} from the other side.");
                }
            }

            double top = slots.Min(slot => slot.Y);
            double bottom = slots.Max(slot => slot.Y);
            double span = bottom - top;
            double centreY = (top + bottom) / 2;
            GateSymbol shape = LogicChipGeometry.ShapeOf(unit.Function);

            double height = shape switch
            {
                GateSymbol.Buffer => LogicChipGeometry.BufferHeight,
                GateSymbol.Mux => LogicChipGeometry.MuxHeightFor(span, LogicChipGeometry.GateWidthOf(unit)),
                GateSymbol.And => Math.Clamp(span + 16, 28, 56),
                _ => Math.Max(28, span + 16),
            };

            double width = LogicChipGeometry.GateWidthOf(unit);
            double bubble = unit.Inverted ? LogicChipGeometry.Bubble : 0;
            double backX = edge + LogicChipGeometry.InputRun;
            double backTop = Math.Min(centreY - (height / 2), top - 6);
            double backBottom = Math.Max(centreY + (height / 2), bottom + 6);
            var output = new Point(backX + width + bubble, centreY);
            var terminals = new List<LogicTerminalAt>();

            foreach ((LogicTerminal terminal, double y, bool isWired) in slots)
            {
                var end = new Point(backX, y);
                terminals.Add(new LogicTerminalAt(terminal.Pin, end));

                if (isWired)
                {
                    part.Wire(terminal.Pin, new Point(edge, y), end);
                }
                else
                {
                    var stub = new Point(backX - LogicChipGeometry.InputStub, y);
                    part.Wire(terminal.Pin, stub, end);
                    part.Badges.Add(new LogicPinBadge(terminal.Pin, LogicChipGeometry.BadgeEndingAt(stub), stub));
                }
            }

            // The output: back to its pin beside the inputs, or across to the other side.
            double outY = Y(unit.Output);
            terminals.Add(new LogicTerminalAt(unit.Output, output));

            if (LogicChipGeometry.IsLeft(chip, unit.Output) == left)
            {
                double bend = output.X + LogicChipGeometry.OutputRun;
                part.Wire(unit.Output, output, new Point(bend, centreY), new Point(bend, outY), new Point(edge, outY));
            }
            else
            {
                double bend = output.X + LogicChipGeometry.JogRun;

                if (Math.Abs(outY - centreY) < 0.01)
                {
                    part.Wire(unit.Output, output, new Point(farEdge, outY));
                }
                else
                {
                    part.Wire(unit.Output, output, new Point(bend, centreY), new Point(bend, outY), new Point(farEdge, outY));
                }
            }

            // The selects and enable on the top edge (or the bottom edge, for one from below).
            double TopEdgeY(double x) => shape == GateSymbol.Buffer
                ? centreY - (height / 2) + ((x - backX) / width * (height / 2))
                : centreY - (height / 2) + ((x - backX) / width * LogicChipGeometry.MuxTaper(height));

            var controls = new List<LogicGateControl>();
            Rect bounds = new(backX, Math.Min(backTop, centreY - (height / 2)), width + bubble, Math.Max(backBottom, centreY + (height / 2)) - Math.Min(backTop, centreY - (height / 2)));

            foreach ((LogicTerminal control, ControlWay way, double offset) in LogicChipGeometry.ControlsOf(chip, unit, left))
            {
                double x = backX + offset;
                bool hasBubble = control.ActiveLow;
                double ring = hasBubble ? LogicChipGeometry.Bubble : 0;
                bool fromBelow = way == ControlWay.Below;
                double edgeY = fromBelow ? (2 * centreY) - TopEdgeY(x) : TopEdgeY(x);
                var edgePoint = new Point(x, edgeY);
                var terminal = new Point(x, fromBelow ? edgeY + ring : edgeY - ring);

                controls.Add(new LogicGateControl(terminal, edgePoint, hasBubble));
                terminals.Add(new LogicTerminalAt(control.Pin, terminal));
                bounds = bounds.Union(new Rect(x - (ring / 2), Math.Min(terminal.Y, edgeY), ring, Math.Abs(terminal.Y - edgeY)));

                if (way != ControlWay.Badge)
                {
                    double rowY = Y(control.Pin);
                    part.Wire(control.Pin, new Point(edge, rowY), new Point(x, rowY), terminal);
                }
                else if (control.Role == TerminalRole.Select)
                {
                    // Up from the top edge, and back to its badge beyond the gate's back.
                    double turn = terminal.Y - LogicChipGeometry.SelectStub;
                    var stub = new Point(backX - BadgeClear, turn);

                    part.Wire(control.Pin, stub, new Point(x, turn), terminal);
                    part.Badges.Add(new LogicPinBadge(control.Pin, LogicChipGeometry.BadgeEndingAt(stub), stub));
                }
                else
                {
                    // Up from above its bubble, and on to its badge beyond the gate's front.
                    double turn = terminal.Y - LogicChipGeometry.EnableStub;
                    var stub = new Point(backX + width + bubble + BadgeClear, turn);

                    part.Wire(control.Pin, stub, new Point(x, turn), terminal);
                    part.Badges.Add(new LogicPinBadge(control.Pin, LogicChipGeometry.BadgeStartingAt(stub), stub));
                }
            }

            List<string> marks = shape == GateSymbol.Mux
                ? [.. slots.Select(slot => slot.Terminal.Index.ToString(System.Globalization.CultureInfo.InvariantCulture))]
                : [.. slots.Select(_ => string.Empty)];

            part.Symbols.Add(new LogicGateShape(
                shape,
                unit.Inverted,
                chip.Outputs,
                chip.Schmitt,
                GateFacing.Right,
                bounds,
                backX,
                width,
                height,
                centreY,
                backTop,
                backBottom,
                [.. slots.Select(slot => new Point(backX, slot.Y))],
                marks,
                output,
                controls,
                [unit],
                terminals));
        }

        // ###########################################################################################
        // A gate placed by a layout of its own (a timer's parts, a buffer of the 244): its shape's back at `backX`, pointing
        // right, about `centreY`, `height` tall, its inputs at the heights given, each with its pin and mark - nothing wired.
        // ###########################################################################################
        private static LogicGateShape GateAt(
            GateSymbol shape, bool inverted, LogicOutputStage stage, double backX, double centreY, double height, double width,
            IReadOnlyList<(int Pin, double Y, string Mark)> inputs, int outputPin, IReadOnlyList<LogicGateControl> controls,
            IReadOnlyList<LogicTerminalAt> controlTerminals, LogicUnit unit)
        {
            double bubble = inverted ? LogicChipGeometry.Bubble : 0;
            var output = new Point(backX + width + bubble, centreY);
            var bounds = new Rect(backX, centreY - (height / 2), width + bubble, height);

            foreach (LogicGateControl control in controls)
            {
                bounds = bounds.Union(new Rect(control.Edge, control.Terminal).Inflate(new Thickness(control.Bubble ? LogicChipGeometry.Bubble / 2 : 0, 0)));
            }

            return new LogicGateShape(
                shape,
                inverted,
                stage,
                Schmitt: false,
                GateFacing.Right,
                bounds,
                backX,
                width,
                height,
                centreY,
                centreY - (height / 2),
                centreY + (height / 2),
                [.. inputs.Select(input => new Point(backX, input.Y))],
                [.. inputs.Select(input => input.Mark)],
                output,
                controls,
                [unit],
                [.. inputs.Select(input => new LogicTerminalAt(input.Pin, new Point(backX, input.Y))), new LogicTerminalAt(outputPin, output), .. controlTerminals]);
        }
    }
}
