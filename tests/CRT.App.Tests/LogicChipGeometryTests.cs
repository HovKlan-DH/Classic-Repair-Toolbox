using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A CHIP'S INSIDES AS DRAWN (LogicChipGeometry; owner request, 2026-11-16). The connections are READ BACK OFF THE DRAWING -
// where each wire starts and ends: on a pin at the case's edge, on a symbol's terminal, or on a pin's badge - and held
// against the chip's units, which LogicChipsTests hold against the datasheets. So every chip CRT knows is drawn as its
// datasheet says, pin by pin, and a drawing that puts a wire on the wrong terminal fails here. Then that it can be read: no
// wire touches another or runs through a symbol, no badge lies on a wire or a symbol, no symbol overlaps another,
// everything stays inside the case - and nothing is drawn at a pin no unit uses (owner request, the same day: no "GND" or
// "VCC" inside the case).
// ###########################################################################################
public sealed class LogicChipGeometryTests
{
    private const double Near = 0.01;

    private static bool Same(Point one, Point other) => Math.Abs(one.X - other.X) < Near && Math.Abs(one.Y - other.Y) < Near;

    // Where a wire from `pin` meets the case's edge - level with its square, on the body's side (LogicChipGeometry's until
    // 2026-11-17, when nothing but these tests asked for it).
    private static Point EdgePointOf(PinoutLayout layout, int pin)
    {
        PinoutPin found = layout.Pins[pin - 1];

        return new Point(found.Side == PinSide.Left ? layout.Body.Left : layout.Body.Right, found.Lead.Center.Y);
    }

    public static TheoryData<string> Chips()
    {
        var data = new TheoryData<string>();

        foreach (LogicChip chip in LogicChips.All)
        {
            data.Add(chip.Name);
        }

        return data;
    }

    private static LogicChip ChipNamed(string name) => LogicChips.All.Single(chip => chip.Name == name);

    private static List<(int Wire, Point A, Point B)> SegmentsOf(LogicInside inside) =>
        inside.Wires
            .SelectMany((wire, index) => wire.Points.Zip(wire.Points.Skip(1), (a, b) => (index, a, b)))
            .ToList();

    // ###########################################################################################
    // Every wire end is a pin at the case's edge, a symbol's terminal, or a pin's badge - its stub meeting the badge on the
    // middle of a side - and a symbol passes on a pin's net between its own terminals of that pin (a column's control
    // section, on to the other column). The nets those ends make are exactly the units' pins: every terminal reaches its
    // own pin and no other - one inside a unit (pin 0) reaches none - and every unit is drawn, each of its pins reached
    // through the symbols standing for it.
    // ###########################################################################################
    [Theory]
    [MemberData(nameof(Chips))]
    public void Every_wire_runs_from_the_pin_the_datasheet_says_to_its_symbol(string name)
    {
        LogicChip chip = ChipNamed(name);
        LogicInside inside = LogicChipGeometry.For(chip);
        PinoutLayout layout = inside.Case;

        var parent = new Dictionary<string, string>();

        string Find(string node)
        {
            parent.TryAdd(node, node);

            while (parent[node] != node)
            {
                node = parent[node];
            }

            return node;
        }

        void Join(string one, string other) => parent[Find(one)] = Find(other);

        var terminals = new List<(Point At, string Node, int Pin)>();

        for (int symbol = 0; symbol < inside.Symbols.Count; symbol++)
        {
            IReadOnlyList<LogicTerminalAt> own = inside.Symbols[symbol].Terminals;

            for (int index = 0; index < own.Count; index++)
            {
                terminals.Add((own[index].At, $"symbol{symbol}:{index}", own[index].Pin));
                Find($"symbol{symbol}:{index}");
            }

            foreach (IGrouping<int, int> samePin in Enumerable.Range(0, own.Count).Where(index => own[index].Pin != 0).GroupBy(index => own[index].Pin))
            {
                foreach (int index in samePin.Skip(1))
                {
                    Join($"symbol{symbol}:{index}", $"symbol{symbol}:{samePin.First()}");
                }
            }
        }

        List<string> NodesOf(Point end)
        {
            for (int pin = 1; pin <= chip.PinCount; pin++)
            {
                if (Same(end, EdgePointOf(layout, pin)))
                {
                    return [$"pin{pin}"];
                }
            }

            List<string> found = terminals.Where(terminal => Same(terminal.At, end)).Select(terminal => terminal.Node).ToList();

            if (found.Count > 0)
            {
                return found;
            }

            LogicPinBadge? badge = inside.Badges.FirstOrDefault(candidate => Same(candidate.Stub, end));
            Assert.True(badge is not null, $"{name}: a wire ends at {end} on nothing");

            // On the middle of the badge's left or right side.
            Rect area = badge!.Area;
            Assert.True(Math.Abs(end.Y - area.Center.Y) < Near && (Math.Abs(end.X - area.Left) < Near || Math.Abs(end.X - area.Right) < Near), $"{name}: pin {badge.Pin}'s badge");

            return [$"pin{badge.Pin}"];
        }

        foreach (LogicWire wire in inside.Wires)
        {
            List<string> ends = [.. NodesOf(wire.Points[0]), .. NodesOf(wire.Points[^1])];

            foreach (string end in ends.Skip(1))
            {
                Join(end, ends[0]);
            }
        }

        List<int> Reached(string node) =>
            Enumerable.Range(1, chip.PinCount).Where(pin => parent.ContainsKey($"pin{pin}") && Find($"pin{pin}") == Find(node)).ToList();

        foreach ((Point at, string node, int pin) in terminals)
        {
            List<int> reached = Reached(node);
            Assert.True(pin == 0 ? reached.Count == 0 : reached.SequenceEqual([pin]), $"{name}: the terminal at {at} reaches pins [{string.Join(", ", reached)}], the datasheet says {pin}");
        }

        // Every unit drawn, every pin of it reached through the symbols standing for it - and no symbol for a unit the
        // chip does not have.
        foreach (LogicUnit unit in chip.Units)
        {
            List<ILogicSymbol> standing = inside.Symbols.Where(symbol => symbol.Units.Any(candidate => ReferenceEquals(candidate, unit))).ToList();
            Assert.NotEmpty(standing);
            Assert.Equal(unit.Pins.Order(), standing.SelectMany(symbol => symbol.Terminals).Select(terminal => terminal.Pin).Where(pin => pin != 0).Distinct().Order());
        }

        Assert.All(inside.Symbols.SelectMany(symbol => symbol.Units), unit => Assert.Contains(chip.Units, candidate => ReferenceEquals(candidate, unit)));
    }

    // ###########################################################################################
    // NOTHING AT A PIN NO UNIT USES (owner request, 2026-11-16: "I do not think this text inside the IC is needed (e.g.
    // "GND" and "VCC")"): the supply, the ground and a pin not connected have no wire and no badge inside the case.
    // ###########################################################################################
    [Theory]
    [MemberData(nameof(Chips))]
    public void Nothing_is_drawn_inside_at_a_pin_no_unit_uses(string name)
    {
        LogicChip chip = ChipNamed(name);
        LogicInside inside = LogicChipGeometry.For(chip);
        HashSet<int> used = chip.Units.SelectMany(unit => unit.Pins).ToHashSet();

        for (int pin = 1; pin <= chip.PinCount; pin++)
        {
            if (!used.Contains(pin))
            {
                Point edge = EdgePointOf(inside.Case, pin);

                Assert.DoesNotContain(inside.Wires, wire => wire.Pin == pin || wire.Points.Any(point => Same(point, edge)));
                Assert.DoesNotContain(inside.Badges, badge => badge.Pin == pin);
            }
        }
    }

    // ###########################################################################################
    // No two wires come within 4 pixels of each other - none crosses or touches another - and no wire enters a symbol:
    // they end on its edges. A badge lies 2 pixels or more from every wire but its own stub, and on no symbol and no other
    // badge.
    // ###########################################################################################
    [Theory]
    [MemberData(nameof(Chips))]
    public void No_wire_crosses_another_and_nothing_lies_on_a_symbol_or_a_badge(string name)
    {
        LogicInside inside = LogicChipGeometry.For(ChipNamed(name));
        List<(int Wire, Point A, Point B)> segments = SegmentsOf(inside);

        for (int one = 0; one < segments.Count; one++)
        {
            for (int other = one + 1; other < segments.Count; other++)
            {
                if (segments[one].Wire != segments[other].Wire)
                {
                    double apart = Distance(segments[one].A, segments[one].B, segments[other].A, segments[other].B);
                    Assert.True(apart >= 4, $"{name}: wires {segments[one].Wire} and {segments[other].Wire} are {apart:0.0} apart");
                }
            }
        }

        foreach (ILogicSymbol symbol in inside.Symbols)
        {
            // A gate's core is its body - its selects' and enable's bubbles stand on its edge, where their wires end.
            Rect core = (symbol is LogicGateShape gate
                ? new Rect(Math.Min(gate.BackX, gate.Output.X), gate.CentreY - (gate.BodyHeight / 2), Math.Abs(gate.Output.X - gate.BackX), gate.BodyHeight)
                : symbol.Bounds).Deflate(new Thickness(1, 4));

            Assert.All(segments, segment => Assert.False(
                Crosses(core, segment.A, segment.B),
                $"{name}: wire {segment.Wire} runs through the symbol at {symbol.Bounds}"));
            Assert.All(inside.Badges, badge => Assert.False(badge.Area.Intersects(symbol.Bounds), $"{name}: pin {badge.Pin}'s badge lies on a symbol"));
        }

        foreach (LogicPinBadge badge in inside.Badges)
        {
            int own = inside.Wires.ToList().FindIndex(wire => wire.Points.Any(point => Same(point, badge.Stub)));

            Assert.All(segments.Where(segment => segment.Wire != own), segment => Assert.True(
                DistanceToRect(badge.Area, segment.A, segment.B) >= 2,
                $"{name}: pin {badge.Pin}'s badge is on wire {segment.Wire}"));

            Assert.All(inside.Badges.Where(other => !ReferenceEquals(other, badge)), other => Assert.False(other.Area.Intersects(badge.Area), $"{name}: badges overlap"));
        }
    }

    // ###########################################################################################
    // Every symbol, wire and badge inside the case, WITH ROOM TO SPARE (owner request, 2026-11-16: "a little vertical/empty
    // spacing between the top/bottom gate borders and the frame of the IC" - the 373's control sections stood 2 pixels from
    // it): a symbol FrameClear from every edge, a badge BadgeFrameClear, and nothing - symbol, badge or wire - within 4
    // pixels of the notch or the dot by pin 1. No symbol over another - blocks may stand edge to edge.
    // ###########################################################################################
    [Theory]
    [MemberData(nameof(Chips))]
    public void Everything_stays_inside_the_case_with_room_to_spare_and_no_symbol_overlaps_another(string name)
    {
        LogicInside inside = LogicChipGeometry.For(ChipNamed(name));
        Rect body = inside.Case.Body;
        var notch = new Point(body.Center.X, body.Top);
        var dot = new Point(body.Left + LogicChipGeometry.FirstPinDot, body.Top + LogicChipGeometry.FirstPinDot);

        Assert.All(inside.Symbols, symbol => Assert.True(body.Deflate(LogicChipGeometry.FrameClear).Contains(symbol.Bounds), $"{name}: {symbol.Bounds} is near the case's edge"));
        Assert.All(inside.Wires.SelectMany(wire => wire.Points), point => Assert.True(body.Contains(point), $"{name}: {point}"));
        Assert.All(inside.Badges, badge => Assert.True(body.Deflate(LogicChipGeometry.BadgeFrameClear).Contains(badge.Area), $"{name}: pin {badge.Pin}'s badge is near the case's edge"));

        foreach (Rect area in inside.Symbols.Select(symbol => symbol.Bounds).Concat(inside.Badges.Select(badge => badge.Area)))
        {
            Assert.True(DistanceToRect(area, notch, notch) >= LogicChipGeometry.NotchRadius + 4, $"{name}: {area} is on the notch");
            Assert.True(DistanceToRect(area, dot, dot) >= LogicChipGeometry.DotRadius + 4, $"{name}: {area} is on pin 1's dot");
        }

        foreach ((int wire, Point a, Point b) in SegmentsOf(inside))
        {
            Assert.True(ToSegment(notch, a, b) >= LogicChipGeometry.NotchRadius + 4, $"{name}: wire {wire} runs past the notch");
            Assert.True(ToSegment(dot, a, b) >= LogicChipGeometry.DotRadius + 4, $"{name}: wire {wire} runs past pin 1's dot");
        }

        // Blocks may stand edge to edge; a gate stands 4 pixels or more from everything else - judged by its outline, as
        // drawn, not its bounds: a timer's comparator and output stage are triangles that nest.
        for (int one = 0; one < inside.Symbols.Count; one++)
        {
            for (int other = one + 1; other < inside.Symbols.Count; other++)
            {
                double least = inside.Symbols[one] is LogicGateShape || inside.Symbols[other] is LogicGateShape ? 4 : 0;

                foreach (Point[] part in PartsOf(inside.Symbols[one]))
                {
                    foreach (Point[] otherPart in PartsOf(inside.Symbols[other]))
                    {
                        double apart = Separation(part, otherPart);
                        Assert.True(apart >= least - 1e-6, $"{name}: symbols {one} and {other} are {apart:0.0} apart");
                    }
                }
            }
        }
    }

    // ###########################################################################################
    // A symbol as the convex shapes it is drawn as: a buffer's or comparator's triangle and a multiplexer's trapezoid, each
    // with its bubbles; a curved gate, a block or a switch by its bounds.
    // ###########################################################################################
    private static List<Point[]> PartsOf(ILogicSymbol symbol)
    {
        static Point[] Box(Rect rect) => [rect.TopLeft, rect.TopRight, rect.BottomRight, rect.BottomLeft];

        if (symbol is not LogicGateShape { Shape: GateSymbol.Buffer or GateSymbol.Comparator or GateSymbol.Mux } gate)
        {
            return [Box(symbol.Bounds)];
        }

        double facing = gate.Facing == GateFacing.Right ? 1 : -1;
        double half = gate.BodyHeight / 2;
        double taper = LogicChipGeometry.MuxTaper(gate.BodyHeight);
        double ring = LogicChipGeometry.Bubble;
        Point At(double x, double y) => new(gate.BackX + (facing * x), gate.CentreY + y);

        List<Point[]> parts =
        [
            gate.Shape == GateSymbol.Mux
                ? [At(0, -half), At(gate.BodyWidth, -half + taper), At(gate.BodyWidth, half - taper), At(0, half)]
                : [At(0, -half), At(gate.BodyWidth, 0), At(0, half)],
        ];

        if (gate.Inverted)
        {
            parts.Add(Box(new Rect(At(gate.BodyWidth, -ring / 2), At(gate.BodyWidth + ring, ring / 2)).Normalize()));
        }

        parts.AddRange(gate.Controls.Where(control => control.Bubble).Select(control => Box(new Rect(control.Edge, control.Terminal).Normalize().Inflate(new Thickness(ring / 2, 0)))));

        return parts;
    }

    // ###########################################################################################
    // How far apart two convex shapes are along the axis that parts them most (the separating axis theorem): negative when
    // they overlap, 0 when they touch, and never more than the true distance - so a gap it gives is there.
    // ###########################################################################################
    private static double Separation(Point[] one, Point[] other)
    {
        double best = double.NegativeInfinity;

        foreach (Point[] shape in new[] { one, other })
        {
            for (int index = 0; index < shape.Length; index++)
            {
                Vector edge = shape[(index + 1) % shape.Length] - shape[index];
                double length = Math.Sqrt((edge.X * edge.X) + (edge.Y * edge.Y));

                if (length == 0)
                {
                    continue;
                }

                var axis = new Vector(-edge.Y / length, edge.X / length);
                double[] a = [.. one.Select(point => (point.X * axis.X) + (point.Y * axis.Y))];
                double[] b = [.. other.Select(point => (point.X * axis.X) + (point.Y * axis.Y))];

                best = Math.Max(best, Math.Max(b.Min() - a.Max(), a.Min() - b.Max()));
            }
        }

        return best;
    }

    // ###########################################################################################
    // A gate stands on the side most of its inputs are on and points into the case. The 7410's first gate takes 1C, from the
    // other side, by pin 13's badge; a multiplexer's select or enable comes from its pin when the pin is just above it on
    // its side - the 257's first takes S, its last G - and from a badge otherwise, nothing drawn at those pins inside the
    // case.
    // ###########################################################################################
    [Fact]
    public void A_gate_points_away_from_its_inputs_and_an_input_from_elsewhere_has_its_pins_badge()
    {
        LogicInside quad = LogicChipGeometry.For(ChipNamed("7408"));

        Assert.Equal(
            [GateFacing.Right, GateFacing.Right, GateFacing.Left, GateFacing.Left],
            quad.Symbols.OfType<LogicGateShape>().Select(gate => gate.Facing));
        Assert.Empty(quad.Badges);

        LogicInside triple = LogicChipGeometry.For(ChipNamed("7410"));
        Assert.Equal([13], triple.Badges.Select(badge => badge.Pin));
        Assert.DoesNotContain(triple.Wires, wire => wire.Points.Any(point => Same(point, EdgePointOf(triple.Case, 13))));

        LogicInside mux = LogicChipGeometry.For(ChipNamed("74257"));
        Assert.Equal(3, mux.Badges.Count(badge => badge.Pin == 1));
        Assert.Equal(3, mux.Badges.Count(badge => badge.Pin == 15));
        Assert.Equal(6, mux.Badges.Count);
        Assert.Contains(mux.Wires, wire => wire.Pin == 1 && Same(wire.Points[0], EdgePointOf(mux.Case, 1)));
        Assert.Contains(mux.Wires, wire => wire.Pin == 15 && Same(wire.Points[0], EdgePointOf(mux.Case, 15)));
    }

    // ###########################################################################################
    // WHAT TAKES A BADGE (owner request, 2026-11-16: "OK, now go ahead and make the rest"): only what no wire can reach - the
    // 153's select from the other side, each 244 buffer's group enable. The 125's enables, the 4066's controls, the flip-flops'
    // and latches' shared controls, the decoders, the counter, the 245 and the timers are wired from their pins.
    // ###########################################################################################
    [Theory]
    [InlineData("74125", "")]
    [InlineData("74153", "14 2")]
    [InlineData("74244", "1 1 1 1 19 19 19 19")]
    [InlineData("74240", "1 1 1 1 19 19 19 19")]
    [InlineData("74245", "")]
    [InlineData("4066", "")]
    [InlineData("7474", "")]
    [InlineData("74175", "")]
    [InlineData("74273", "")]
    [InlineData("74373", "")]
    [InlineData("74138", "")]
    [InlineData("74139", "")]
    [InlineData("74145", "")]
    [InlineData("74193", "")]
    [InlineData("555", "")]
    [InlineData("556", "")]
    public void Only_what_no_wire_can_reach_takes_a_badge(string name, string badges)
    {
        LogicInside inside = LogicChipGeometry.For(ChipNamed(name));

        Assert.Equal(
            badges.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).Order(),
            inside.Badges.Select(badge => badge.Pin).Order());
    }

    // ###########################################################################################
    // The shared controls of an array (the 373's OE and LE) come into a common control section at the top and bottom of
    // each column: from the pin into the section on its side, and on across into the other's - one wire each way, so every
    // latch is seen to take them.
    // ###########################################################################################
    [Fact]
    public void A_column_array_takes_its_shared_controls_through_a_section_at_each_end_of_each_column()
    {
        LogicChip chip = ChipNamed("74373");
        LogicInside inside = LogicChipGeometry.For(chip);
        List<LogicBlockShape> blocks = inside.Symbols.OfType<LogicBlockShape>().ToList();

        // Eight latches and four sections, each section standing for its column's four latches.
        Assert.Equal(8, blocks.Count(block => block.Units.Count == 1));
        List<LogicBlockShape> sections = blocks.Where(block => block.Units.Count == 4).ToList();
        Assert.Equal(4, sections.Count);

        foreach (int pin in new[] { 1, 11 })
        {
            Assert.Equal(2, sections.Count(section => section.Ports.Any(port => port.Pin == pin)));
            Assert.Equal(2, inside.Wires.Count(wire => wire.Pin == pin));
        }

        // OE in at the top, LE at the bottom - each section over its control's row, as an element is over its own.
        double oe = inside.Case.Pins[0].Lead.Center.Y;
        double le = inside.Case.Pins[10].Lead.Center.Y;

        Assert.All(sections.Where(section => section.Ports.Any(port => port.Pin == 1)), section =>
            Assert.True(section.Bounds.Top <= blocks.Min(block => block.Bounds.Top) && section.Bounds.Top < oe && section.Bounds.Bottom > oe));
        Assert.All(sections.Where(section => section.Ports.Any(port => port.Pin == 11)), section =>
            Assert.True(section.Bounds.Bottom >= blocks.Max(block => block.Bounds.Bottom) && section.Bounds.Top < le && section.Bounds.Bottom > le));
    }

    // ###########################################################################################
    // Every mark written in a block is plain ASCII, as every user-visible text is (CLAUDE.md; code review, 2026-11-17: the
    // 74193's count-down clock was a Unicode minus) - and that clock is marked "-".
    // ###########################################################################################
    [Fact]
    public void Every_mark_written_in_a_block_is_plain_ASCII()
    {
        foreach (LogicChip chip in LogicChips.All)
        {
            foreach (LogicBlockPort port in LogicChipGeometry.For(chip).Symbols.OfType<LogicBlockShape>().SelectMany(block => block.Ports))
            {
                Assert.True(port.Mark.All(char.IsAscii), $"{chip.Name} pin {port.Pin}: [{port.Mark}]");
            }
        }

        LogicChip counter = ChipNamed("74193");
        int down = counter.Units.SelectMany(unit => unit.Of(TerminalRole.CountDown)).Single().Pin;
        Assert.Contains(
            LogicChipGeometry.For(counter).Symbols.OfType<LogicBlockShape>().SelectMany(block => block.Ports),
            port => port.Pin == down && port.Mark == "-");
    }

    // ###########################################################################################
    // A 555's insides: two comparators, the flip-flop, the output stage and the inverted, open-collector discharge - each
    // joined to the flip-flop inside the timer, the third of the supply written at its comparator, not wired.
    // ###########################################################################################
    [Fact]
    public void A_timer_is_drawn_as_its_comparators_flip_flop_and_output_stages()
    {
        LogicInside inside = LogicChipGeometry.For(ChipNamed("555"));
        List<LogicGateShape> gates = inside.Symbols.OfType<LogicGateShape>().ToList();

        Assert.Equal(2, gates.Count(gate => gate.Shape == GateSymbol.Comparator));
        Assert.Single(gates, gate => gate.Shape == GateSymbol.Buffer && !gate.Inverted && gate.Stage == LogicOutputStage.TotemPole && gate.Terminals.Any(terminal => terminal.Pin == 3));
        Assert.Single(gates, gate => gate.Shape == GateSymbol.Buffer && gate.Inverted && gate.Stage == LogicOutputStage.OpenCollector && gate.Terminals.Any(terminal => terminal.Pin == 7));
        Assert.Single(inside.Symbols.OfType<LogicBlockShape>());
        Assert.Equal(4, inside.Wires.Count(wire => wire.Pin == 0));
        Assert.Contains(gates, gate => gate.InputTags.Contains(LogicChipGeometry.Third));
    }

    // ###########################################################################################
    // *** EVERY MARK INSIDE A GATE KEEPS CLEAR OF ITS OUTLINE (owner request, 2026-11-17: "the +/- is too close to its
    // borders ... There could be others with same issue"). *** A 556's "-" lay on its comparator's slanted edge, the 257's
    // "0" a pixel under its multiplexer's top edge, the 7414's Schmitt mark and the 7406's open-collector mark ran into their
    // inverters' sides. Every mark's ink - a sign's arm tips, a number's, a stage mark's area - is MarkClear or more from the
    // outline's ink of every gate with straight edges (the strokes' halves counted), measured here from the outline as the
    // painter draws it: a buffer's or comparator's triangle, a multiplexer's trapezoid. A curved gate's mark stands in its
    // middle; it is held inside its body. And no two marks of a gate meet, and a comparator's signs are "+" and "-" as strokes.
    // ###########################################################################################
    [Theory]
    [MemberData(nameof(Chips))]
    public void Every_mark_inside_a_gate_keeps_clear_of_its_outline(string name)
    {
        double needed = LogicChipGeometry.MarkClear + (LogicChipGeometry.GateOutline / 2) + (LogicChipGeometry.MarkStroke / 2);

        foreach (LogicGateShape gate in LogicChipGeometry.For(ChipNamed(name)).Symbols.OfType<LogicGateShape>())
        {
            double facing = gate.Facing == GateFacing.Right ? 1 : -1;
            double half = gate.BodyHeight / 2;
            double taper = LogicChipGeometry.MuxTaper(gate.BodyHeight);
            Point At(double x, double y) => new(gate.BackX + (facing * x), gate.CentreY + y);

            Point[]? outline = gate.Shape switch
            {
                GateSymbol.Buffer or GateSymbol.Comparator => [At(0, -half), At(gate.BodyWidth, 0), At(0, half)],
                GateSymbol.Mux => [At(0, -half), At(gate.BodyWidth, -half + taper), At(gate.BodyWidth, half - taper), At(0, half)],
                _ => null,
            };

            IReadOnlyList<LogicGateMark> marks = LogicChipGeometry.MarksOf(gate);

            foreach (LogicGateMark mark in marks)
            {
                Assert.All(mark.Ink, point => Assert.True(mark.Area.Inflate(Near).Contains(point), $"{name}: a {mark.Kind}'s ink outside its area"));

                if (outline is null)
                {
                    var body = new Rect(Math.Min(gate.BackX, gate.BackX + (facing * gate.BodyWidth)), gate.CentreY - half, gate.BodyWidth, gate.BodyHeight);
                    Assert.True(body.Deflate(LogicChipGeometry.MarkInset).Contains(mark.Area), $"{name}: a {mark.Kind} mark near its {gate.Shape}'s edge");
                    continue;
                }

                foreach (Point point in mark.Ink)
                {
                    Assert.True(Inside(outline, point), $"{name}: a {mark.Kind} mark outside its {gate.Shape}");

                    double room = outline.Select((corner, index) => ToSegment(point, corner, outline[(index + 1) % outline.Length])).Min();
                    Assert.True(room >= needed - 1e-6, $"{name}: a {mark.Kind} mark {room - needed + LogicChipGeometry.MarkClear:0.00} from its {gate.Shape}'s outline");
                }
            }

            for (int one = 0; one < marks.Count; one++)
            {
                for (int other = one + 1; other < marks.Count; other++)
                {
                    Assert.False(marks[one].Area.Intersects(marks[other].Area), $"{name}: a {marks[one].Kind} and a {marks[other].Kind} meet");
                }
            }

            if (gate.Shape == GateSymbol.Comparator)
            {
                Assert.Equal(
                    gate.InputMarks.Where(mark => mark.Length > 0).Select(mark => mark == "+" ? GateMarkKind.Plus : GateMarkKind.Minus),
                    marks.Select(mark => mark.Kind));
            }
        }
    }

    // ###########################################################################################
    // A gate's stage or Schmitt mark is drawn as large as fits, down to half its size - and left out only where even that does
    // not fit: the 240's and 244's buffers, half a row apart, whose enables say they are 3-state. Any other chip losing one is
    // a decision, added here.
    // ###########################################################################################
    [Fact]
    public void Only_the_octal_buffers_are_too_small_for_their_stage_mark()
    {
        static bool HasInnerMark(LogicGateShape gate) => gate.Stage != LogicOutputStage.TotemPole || gate.Schmitt;

        static bool DrawsIt(LogicGateShape gate) => LogicChipGeometry.MarksOf(gate).Any(mark => mark.Kind is GateMarkKind.ThreeState or GateMarkKind.OpenCollector or GateMarkKind.Schmitt);

        List<string> without = LogicChips.All
            .Where(chip => LogicChipGeometry.For(chip).Symbols.OfType<LogicGateShape>().Any(gate => HasInnerMark(gate) && !DrawsIt(gate)))
            .Select(chip => chip.Name)
            .ToList();

        Assert.Equal(["74240", "74244"], without);

        // The 7414's Schmitt marks, the 7406's open-collector ones and the 125's 3-state ones are drawn - at most a little smaller.
        foreach ((string chip, GateMarkKind kind, double width) in new[] { ("7414", GateMarkKind.Schmitt, 10.0), ("7406", GateMarkKind.OpenCollector, 6.8), ("74125", GateMarkKind.ThreeState, 7.0) })
        {
            Assert.All(LogicChipGeometry.For(ChipNamed(chip)).Symbols.OfType<LogicGateShape>(), gate =>
            {
                LogicGateMark mark = Assert.Single(LogicChipGeometry.MarksOf(gate), candidate => candidate.Kind == kind);
                Assert.True(mark.Area.Width >= 0.8 * width, $"{chip}: its {kind} mark is {mark.Area.Width:0.0} wide");
            });
        }
    }

    // ###########################################################################################
    // "1/3" AT THE MARKS' SIZE (owner request, 2026-11-17: 'this "1/3" text could have same text size as the other texts inside
    // this IC'): written in plain ASCII in an area outside its comparator's back that holds it, as clear of the case's edge as
    // a badge, and on no wire, symbol or badge.
    // ###########################################################################################
    [Theory]
    [InlineData("555")]
    [InlineData("556")]
    public void A_timers_third_is_written_where_nothing_else_is(string name)
    {
        LogicInside inside = LogicChipGeometry.For(ChipNamed(name));
        Rect body = inside.Case.Body;
        List<Rect> areas = inside.Symbols.OfType<LogicGateShape>()
            .SelectMany(gate => gate.InputTags.Select((tag, index) => (tag, index)).Where(tagged => tagged.tag.Length > 0).Select(tagged => LogicChipGeometry.TagArea(gate, tagged.index)))
            .ToList();

        Assert.Equal(name == "556" ? 2 : 1, areas.Count);
        Assert.Equal("1/3", LogicChipGeometry.Third);

        foreach (Rect area in areas)
        {
            Assert.True(body.Deflate(LogicChipGeometry.BadgeFrameClear).Contains(area), $"{name}: the third's {area} is near the case's edge");
            Assert.All(inside.Symbols, symbol => Assert.False(symbol.Bounds.Intersects(area), $"{name}: the third is on a symbol"));
            Assert.All(inside.Badges, badge => Assert.False(badge.Area.Intersects(area), $"{name}: the third is on a badge"));
            Assert.All(SegmentsOf(inside), segment => Assert.True(DistanceToRect(area, segment.A, segment.B) >= 2, $"{name}: the third is on wire {segment.Wire}"));
        }
    }

    // ###########################################################################################
    // EVERY DIP IN THE SAME CASE (owner request, 2026-11-17: "unify ALL pinout IC's so they look the same, meaning that IC's
    // where it has no data (like e.g. the CPU or VIC-II etc) will look like the new ones"): one CRT knows nothing inside of is
    // that case empty, its body a real one's proportion - three pitches (300 mil) wide, six (600 mil) for a package named
    // "wide" or of 24 pins or more - and no case for what is no DIP CRT can count.
    // ###########################################################################################
    [Theory]
    [InlineData("DIP-8", 8, 3)]
    [InlineData("DIP-14", 14, 3)]
    [InlineData("DIP-20", 20, 3)]
    [InlineData("DIP-20 wide", 20, 6)]
    [InlineData("DIP-24", 24, 6)]
    [InlineData("DIP-28", 28, 6)]
    [InlineData("DIP-40", 40, 6)]
    [InlineData("DIP-64", 64, 6)]
    // A surface-mounted chip numbered as a DIP is drawn in its case (owner request, 2026-11-18).
    [InlineData("SOIC-14", 14, 3)]
    [InlineData("SSOP-28", 28, 6)]
    public void Any_other_DIP_is_the_same_case_empty_and_as_wide_as_a_real_one(string package, int pins, int pitches)
    {
        LogicInside inside = Assert.IsType<LogicInside>(LogicChipGeometry.EmptyCaseFor(package));

        Assert.Equal(pins, inside.Case.Pins.Count);
        Assert.Equal(pitches * LogicChipGeometry.Pitch, inside.Case.Body.Width);
        Assert.Equal(((pins / 2) * LogicChipGeometry.Pitch) + (2 * LogicChipGeometry.EndPadding), inside.Case.Body.Height);
        Assert.Equal(LogicChipGeometry.Pitch, inside.Case.Pitch);
        Assert.Empty(inside.Symbols);
        Assert.Empty(inside.Wires);
        Assert.Empty(inside.Badges);
    }

    [Theory]
    [InlineData("TO-92")]
    [InlineData("SIP")]
    [InlineData("DIP")]
    [InlineData("DIP-7")]
    [InlineData("PLCC-32")]
    [InlineData("SOT-23")]
    [InlineData("")]
    [InlineData(null)]
    public void What_is_no_chips_case_CRT_can_count_has_no_case(string? package) =>
        Assert.Null(LogicChipGeometry.EmptyCaseFor(package));

    // A PLCC or a QFP is the square case of four sides, empty (QuadCaseGeometryTests holds its pins).
    [Theory]
    [InlineData("PLCC-44", 44)]
    [InlineData("PLCC-84", 84)]
    [InlineData("QFP-44", 44)]
    public void A_case_of_four_sides_is_the_square_case_empty(string package, int pins)
    {
        LogicInside inside = Assert.IsType<LogicInside>(LogicChipGeometry.EmptyCaseFor(package));

        Assert.Equal(pins, inside.Case.Pins.Count);
        Assert.True(inside.Case.IsQuad);
        Assert.Equal(inside.Case.Body.Width, inside.Case.Body.Height);
        Assert.Empty(inside.Symbols);
        Assert.Empty(inside.Wires);
        Assert.Empty(inside.Badges);
    }

    // ###########################################################################################
    // A CRYSTAL OSCILLATOR'S CAN (owner request, 2026-11-18 - the Amiga's): its DIP's places with a pin at each corner alone,
    // pin 1 at the top left and the rest where a DIP has them - so the pointer finds each - its body in a real can's
    // proportion, a full-size one 13 by 20.5 mm and a half-size one square.
    // ###########################################################################################
    [Theory]
    [InlineData("Oscillator DIP-14", new[] { 1, 7, 8, 14 }, 13.0 / 20.5)]
    [InlineData("Oscillator DIP-8", new[] { 1, 4, 5, 8 }, 1.0)]
    public void An_oscillators_can_has_its_four_pins_at_its_corners(string package, int[] numbers, double widthToHeight)
    {
        LogicInside inside = Assert.IsType<LogicInside>(LogicChipGeometry.EmptyCaseFor(package));
        PinoutLayout can = inside.Case;
        PinoutLayout places = LogicChipGeometry.Case(numbers[^1], can.Body.Width);

        Assert.True(can.IsCan);
        Assert.False(can.IsQuad);
        Assert.Equal(numbers, can.Pins.Select(pin => pin.Number));
        Assert.Equal(numbers.Select(number => places.Pins[number - 1]), can.Pins);
        Assert.Equal([PinSide.Left, PinSide.Left, PinSide.Right, PinSide.Right], can.Pins.Select(pin => pin.Side));
        Assert.Equal(can.Pins[0].Lead.Center.Y, can.Pins[3].Lead.Center.Y);
        Assert.Equal(can.Pins[1].Lead.Center.Y, can.Pins[2].Lead.Center.Y);
        Assert.Equal(widthToHeight, can.Body.Width / can.Body.Height, 2);
        Assert.Empty(inside.Symbols);

        foreach (PinoutPin pin in can.Pins)
        {
            Assert.Equal(pin.Number, PinoutGeometry.PinAt(can, pin.Lead.Center)?.Number);
        }
    }

    // ###########################################################################################
    // The case: pins down the left and up the right 24 apart, each a square beside the body - so the pointer finds every
    // pin there (PinoutGeometry.PinAt, which goes by the layout's pitch), and the case grows with what is inside it.
    // ###########################################################################################
    [Fact]
    public void The_case_numbers_its_pins_as_a_DIP_and_the_pointer_finds_each_on_its_square()
    {
        PinoutLayout layout = LogicChipGeometry.For(ChipNamed("74257"), nameWidth: 40).Case;

        Assert.Equal(LogicChipGeometry.Pitch, layout.Pitch);
        Assert.Equal(16, layout.Pins.Count);
        Assert.Equal(PinSide.Left, layout.Pins[0].Side);
        Assert.Equal(PinSide.Right, layout.Pins[15].Side);
        Assert.Equal(layout.Pins[0].Lead.Center.Y, layout.Pins[15].Lead.Center.Y);
        Assert.Equal(layout.Pins[7].Lead.Center.Y, layout.Pins[8].Lead.Center.Y);
        Assert.Equal(LogicChipGeometry.Pitch, layout.Pins[1].Lead.Center.Y - layout.Pins[0].Lead.Center.Y);

        Assert.All(layout.Pins, pin =>
        {
            Assert.True(pin.Side == PinSide.Left ? pin.Lead.Right < layout.Body.Left : pin.Lead.Left > layout.Body.Right);
            Assert.Equal(pin.Number, PinoutGeometry.PinAt(layout, pin.Lead.Center)?.Number);
            Assert.Equal(pin.Number, PinoutGeometry.PinAt(layout, pin.Lead.Center + new Point(0, (LogicChipGeometry.Pitch / 2) - 1))?.Number);
        });

        Assert.Equal(new Size(layout.Body.Right + LogicChipGeometry.LeadGap + LogicChipGeometry.Square + 40, layout.Body.Bottom), layout.Size);

        // One gate on one side: the case is narrower, but has some width on its empty side.
        Rect single = LogicChipGeometry.For(ChipNamed("74133")).Case.Body;
        Assert.True(single.Width < LogicChipGeometry.For(ChipNamed("7408")).Case.Body.Width);
        Assert.True(single.Width >= LogicChipGeometry.LeastReach * 2);
    }

    // ---- geometry helpers -------------------------------------------------------------------------

    // Whether a point lies inside a convex outline, on whichever side its corners run.
    private static bool Inside(Point[] outline, Point point)
    {
        static double Cross(Point o, Point p, Point q) => ((p.X - o.X) * (q.Y - o.Y)) - ((p.Y - o.Y) * (q.X - o.X));

        double[] sides = [.. outline.Select((corner, index) => Cross(corner, outline[(index + 1) % outline.Length], point))];
        return sides.All(side => side > 0) || sides.All(side => side < 0);
    }

    private static double Distance(Point a, Point b, Point c, Point d)
    {
        if (Intersect(a, b, c, d))
        {
            return 0;
        }

        return new[] { ToSegment(a, c, d), ToSegment(b, c, d), ToSegment(c, a, b), ToSegment(d, a, b) }.Min();
    }

    private static double ToSegment(Point p, Point a, Point b)
    {
        Vector ab = b - a;
        double length = (ab.X * ab.X) + (ab.Y * ab.Y);
        double t = length == 0 ? 0 : Math.Clamp((((p.X - a.X) * ab.X) + ((p.Y - a.Y) * ab.Y)) / length, 0, 1);
        Point closest = a + (ab * t);

        return Math.Sqrt(((p.X - closest.X) * (p.X - closest.X)) + ((p.Y - closest.Y) * (p.Y - closest.Y)));
    }

    private static bool Intersect(Point a, Point b, Point c, Point d)
    {
        static double Cross(Point o, Point p, Point q) => ((p.X - o.X) * (q.Y - o.Y)) - ((p.Y - o.Y) * (q.X - o.X));

        double d1 = Cross(c, d, a), d2 = Cross(c, d, b), d3 = Cross(a, b, c), d4 = Cross(a, b, d);

        return ((d1 > 0 && d2 < 0) || (d1 < 0 && d2 > 0)) && ((d3 > 0 && d4 < 0) || (d3 < 0 && d4 > 0));
    }

    // Whether the segment from a to b (level or upright - every wire is) passes through the rectangle.
    private static bool Crosses(Rect rect, Point a, Point b)
    {
        double left = Math.Min(a.X, b.X), right = Math.Max(a.X, b.X), top = Math.Min(a.Y, b.Y), bottom = Math.Max(a.Y, b.Y);

        return right > rect.Left && left < rect.Right && bottom > rect.Top && top < rect.Bottom;
    }

    // How far a level or upright segment is from a rectangle - 0 when it touches or crosses it.
    private static double DistanceToRect(Rect rect, Point a, Point b)
    {
        double left = Math.Min(a.X, b.X), right = Math.Max(a.X, b.X), top = Math.Min(a.Y, b.Y), bottom = Math.Max(a.Y, b.Y);
        double dx = Math.Max(0, Math.Max(rect.Left - right, left - rect.Right));
        double dy = Math.Max(0, Math.Max(rect.Top - bottom, top - rect.Bottom));

        return Math.Sqrt((dx * dx) + (dy * dy));
    }
}
