using Avalonia;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE SYMBOLS INSIDE A CHIP, PAINTED (LogicChipGeometry places them) - outlined in the theme's text color on the case's
    // ground, as the datasheets draw them, what is written at a terminal in the muted text color:
    //
    // - A GATE: AND's D, OR's curved back, XOR's second back, a buffer's triangle, a multiplexer's trapezoid with its inputs'
    //   numbers, a comparator's triangle with "+" and "-" - and a reference written outside its back - an inverting output's
    //   bubble, an active-low select's or enable's bubble on its edge, and the marks for an open-collector output, a
    //   Schmitt-trigger input and a 3-state output.
    // - A BLOCK: its rectangle and at each terminal its mark ("D", "Q" - overlined for Q-bar - "EN", a weight), a clock's
    //   wedge, an active-low terminal's bubble and a 3-state or open-collector output's mark; a transceiver's two ways in its
    //   middle.
    // - A SWITCH: its two contacts, its blade drawn open, and the dashed link from its control's wire to the blade.
    // ###########################################################################################
    public sealed partial class PinoutDrawing
    {
        private const double MarkSize = 8.5;

        // The brushes a symbol is painted with: its outline, the case's ground it is filled with, and its marks - and the color
        // of a wire inside a unit, for the stub to a tag.
        private sealed record SymbolInk(IBrush Line, IBrush Fill, IBrush Mark, Color Wire);

        // ---- gates -------------------------------------------------------------------------------------

        // ###########################################################################################
        // One gate, from its own outline pointing right - mirrored for one pointing left: its extended back line, its shape
        // (and the stubs from its curved back to where its inputs' wires end), its bubbles, and its marks.
        // ###########################################################################################
        private void DrawGate(DrawingContext context, LogicGateShape gate, SymbolInk ink)
        {
            double sign = gate.Facing == GateFacing.Right ? 1 : -1;
            double w = gate.BodyWidth;
            double h = gate.BodyHeight;
            Point At(double x, double y) => new(gate.BackX + (sign * x), gate.CentreY + y);

            var outline = new Pen(ink.Line, LogicChipGeometry.GateOutline, lineJoin: PenLineJoin.Round);

            // The datasheets' extended back line, for more inputs than the shape is tall.
            if (gate.BackTop < gate.CentreY - (h / 2) - 0.5 || gate.BackBottom > gate.CentreY + (h / 2) + 0.5)
            {
                context.DrawLine(outline, new Point(gate.BackX, gate.BackTop), new Point(gate.BackX, gate.BackBottom));
            }

            var shape = new StreamGeometry();
            double xorShift = gate.Shape == GateSymbol.Xor ? 5 : 0;
            double orWidth = w - xorShift;

            using (StreamGeometryContext path = shape.Open())
            {
                switch (gate.Shape)
                {
                    case GateSymbol.And:
                    {
                        double r = h / 2;
                        SweepDirection sweep = sign > 0 ? SweepDirection.Clockwise : SweepDirection.CounterClockwise;

                        path.BeginFigure(At(0, -r), isFilled: true);
                        path.LineTo(At(w - r, -r));
                        path.ArcTo(At(w, 0), new Size(r, r), 0, false, sweep);
                        path.ArcTo(At(w - r, r), new Size(r, r), 0, false, sweep);
                        path.LineTo(At(0, r));
                        path.EndFigure(isClosed: true);
                        break;
                    }

                    case GateSymbol.Or:
                    case GateSymbol.Xor:
                        path.BeginFigure(At(xorShift, -h / 2), isFilled: true);
                        path.QuadraticBezierTo(At(xorShift + (orWidth * 0.6), -h / 2), At(w, 0));
                        path.QuadraticBezierTo(At(xorShift + (orWidth * 0.6), h / 2), At(xorShift, h / 2));
                        path.QuadraticBezierTo(At(xorShift + (orWidth * 0.28), 0), At(xorShift, -h / 2));
                        path.EndFigure(isClosed: true);
                        break;

                    case GateSymbol.Buffer:
                    case GateSymbol.Comparator:
                        path.BeginFigure(At(0, -h / 2), isFilled: true);
                        path.LineTo(At(w, 0));
                        path.LineTo(At(0, h / 2));
                        path.EndFigure(isClosed: true);
                        break;

                    default:
                    {
                        double front = (h / 2) - LogicChipGeometry.MuxTaper(h);

                        path.BeginFigure(At(0, -h / 2), isFilled: true);
                        path.LineTo(At(w, -front));
                        path.LineTo(At(w, front));
                        path.LineTo(At(0, h / 2));
                        path.EndFigure(isClosed: true);
                        break;
                    }
                }
            }

            // An OR's or XOR's inputs reach on to its curved back.
            if (gate.Shape is GateSymbol.Or or GateSymbol.Xor)
            {
                foreach (Point input in gate.Inputs)
                {
                    double t = (1 - (2 * (input.Y - gate.CentreY) / h)) / 2;
                    double depth = 2 * t * (1 - t) * (orWidth * 0.28);

                    context.DrawLine(outline, input, At(depth, input.Y - gate.CentreY));
                }
            }

            context.DrawGeometry(ink.Fill, outline, shape);

            // An XOR's second back.
            if (gate.Shape == GateSymbol.Xor)
            {
                var back = new StreamGeometry();

                using (StreamGeometryContext path = back.Open())
                {
                    path.BeginFigure(At(0, h / 2), isFilled: false);
                    path.QuadraticBezierTo(At(orWidth * 0.28, 0), At(0, -h / 2));
                    path.EndFigure(isClosed: false);
                }

                context.DrawGeometry(null, outline, back);
            }

            if (gate.Inverted)
            {
                context.DrawEllipse(ink.Fill, outline, At(w + (LogicChipGeometry.Bubble / 2), 0), LogicChipGeometry.Bubble / 2, LogicChipGeometry.Bubble / 2);
            }

            // An active-low select's or enable's bubble, between the edge and where its wire ends.
            foreach (LogicGateControl control in gate.Controls.Where(control => control.Bubble))
            {
                var centre = new Point((control.Edge.X + control.Terminal.X) / 2, (control.Edge.Y + control.Terminal.Y) / 2);
                context.DrawEllipse(ink.Fill, outline, centre, LogicChipGeometry.Bubble / 2, LogicChipGeometry.Bubble / 2);
            }

            this.DrawGateMarks(context, gate, sign, ink);
        }

        // ###########################################################################################
        // A gate's marks, each filling the area LogicChipGeometry.MarksOf keeps for it clear of the outline: a comparator's
        // "+" and "-" as strokes, a multiplexer's numbers - the digit's ink in the middle of its area - and an open-collector
        // output's, a 3-state output's or a Schmitt-trigger input's mark inside it. A reference outside its back is written at
        // the marks' size at the end of a stub of wire.
        // ###########################################################################################
        private void DrawGateMarks(DrawingContext context, LogicGateShape gate, double sign, SymbolInk ink)
        {
            var markPen = new Pen(ink.Mark, LogicChipGeometry.MarkStroke);
            var linePen = new Pen(ink.Line, LogicChipGeometry.MarkStroke, lineJoin: PenLineJoin.Round);

            foreach (LogicGateMark mark in this.MarksOf(gate))
            {
                Rect area = mark.Area;
                Point centre = area.Center;
                var geometry = new StreamGeometry();

                switch (mark.Kind)
                {
                    case GateMarkKind.Plus:
                    case GateMarkKind.Minus:
                        context.DrawLine(markPen, new Point(area.Left, centre.Y), new Point(area.Right, centre.Y));

                        if (mark.Kind == GateMarkKind.Plus)
                        {
                            context.DrawLine(markPen, new Point(centre.X, area.Top), new Point(centre.X, area.Bottom));
                        }

                        continue;

                    case GateMarkKind.Text:
                    {
                        FormattedText text = PinoutDrawing.TextOf(mark.Text, ink.Mark, MarkSize, bold: false);
                        Rect inked = this.InkedBoundsOf(text, mark.Text);
                        context.DrawText(text, new Point(centre.X - inked.Center.X, centre.Y - inked.Center.Y));
                        continue;
                    }

                    case GateMarkKind.ThreeState:
                        using (StreamGeometryContext path = geometry.Open())
                        {
                            path.BeginFigure(area.TopLeft, isFilled: false);
                            path.LineTo(area.TopRight);
                            path.LineTo(new Point(centre.X, area.Bottom));
                            path.EndFigure(isClosed: true);
                        }

                        break;

                    case GateMarkKind.OpenCollector:
                    {
                        // A diamond over a bar, the diamond as wide as the area.
                        double size = area.Width / 2;
                        double middle = area.Top + size;

                        using (StreamGeometryContext path = geometry.Open())
                        {
                            path.BeginFigure(new Point(area.Left, middle), isFilled: false);
                            path.LineTo(new Point(centre.X, area.Top));
                            path.LineTo(new Point(area.Right, middle));
                            path.LineTo(new Point(centre.X, area.Top + area.Width));
                            path.EndFigure(isClosed: true);
                            path.BeginFigure(area.BottomLeft, isFilled: false);
                            path.LineTo(area.BottomRight);
                            path.EndFigure(isClosed: false);
                        }

                        break;
                    }

                    case GateMarkKind.Schmitt:
                    {
                        // The hysteresis loop: two steps facing each other, a tenth of its width apart from its middle.
                        double u = sign * area.Width / 10;
                        Point At(double x, double y) => new(centre.X + (u * x), centre.Y + (area.Height / 6 * y));

                        using (StreamGeometryContext path = geometry.Open())
                        {
                            path.BeginFigure(At(-5, 3), isFilled: false);
                            path.LineTo(At(1, 3));
                            path.LineTo(At(1, -3));
                            path.EndFigure(isClosed: false);
                            path.BeginFigure(At(-1, 3), isFilled: false);
                            path.LineTo(At(-1, -3));
                            path.LineTo(At(5, -3));
                            path.EndFigure(isClosed: false);
                        }

                        break;
                    }
                }

                context.DrawGeometry(null, linePen, geometry);
            }

            for (int index = 0; index < gate.Inputs.Count; index++)
            {
                string tag = index < gate.InputTags.Count ? gate.InputTags[index] : string.Empty;

                if (tag.Length > 0)
                {
                    Point input = gate.Inputs[index];
                    FormattedText text = PinoutDrawing.TextOf(tag, ink.Mark, MarkSize, bold: false);
                    Rect area = LogicChipGeometry.TagArea(gate, index);
                    var lineEnd = new Point(input.X - (sign * LogicChipGeometry.TagLine), input.Y);
                    double x = sign > 0 ? area.Right - text.Width : area.Left;

                    context.DrawLine(new Pen(new SolidColorBrush(ink.Wire), WireWidth), lineEnd, input);
                    context.DrawText(text, new Point(x, area.Center.Y - (text.Height / 2)));
                }
            }
        }

        // ---- blocks ------------------------------------------------------------------------------------

        // ###########################################################################################
        // A block: its rectangle, and at each terminal its bubble, a clock's wedge, its mark (overlined for Q-bar) and an
        // output's 3-state or open-collector mark beside it; a transceiver's two ways - a triangle each way - in its middle.
        // ###########################################################################################
        private static void DrawBlock(DrawingContext context, LogicBlockShape block, SymbolInk ink)
        {
            var outline = new Pen(ink.Line, 1.2, lineJoin: PenLineJoin.Round);
            var thin = new Pen(ink.Line, 1, lineJoin: PenLineJoin.Round);
            var markPen = new Pen(ink.Mark, 1);
            Rect rect = block.Bounds;

            context.DrawRectangle(ink.Fill, outline, rect);

            foreach (LogicBlockPort port in block.Ports)
            {
                // Where it meets the edge - inside its bubble, when it has one - and which way is inward.
                (double dx, double dy) = port.Edge switch
                {
                    BlockEdge.Left => (1.0, 0.0),
                    BlockEdge.Right => (-1.0, 0.0),
                    BlockEdge.Top => (0.0, 1.0),
                    _ => (0.0, -1.0),
                };

                double ring = port.Bubble ? LogicChipGeometry.Bubble : 0;
                var edge = new Point(port.At.X + (dx * ring), port.At.Y + (dy * ring));

                if (port.Bubble)
                {
                    var centre = new Point(port.At.X + (dx * ring / 2), port.At.Y + (dy * ring / 2));
                    context.DrawEllipse(ink.Fill, outline, centre, ring / 2, ring / 2);
                }

                double inset = 3;

                if (port.Clock)
                {
                    var wedge = new StreamGeometry();

                    using (StreamGeometryContext path = wedge.Open())
                    {
                        path.BeginFigure(new Point(edge.X + (dy * 4), edge.Y + (dx * 4)), isFilled: false);
                        path.LineTo(new Point(edge.X + (dx * 6), edge.Y + (dy * 6)));
                        path.LineTo(new Point(edge.X - (dy * 4), edge.Y - (dx * 4)));
                        path.EndFigure(isClosed: false);
                    }

                    context.DrawGeometry(null, thin, wedge);
                    inset = 8;
                }

                double markEnd = PinoutDrawing.DrawPortMark(context, port.Mark, port.Overline, edge, port.Edge, inset, ink.Mark, markPen);

                if (port.Stage is LogicOutputStage stage)
                {
                    PinoutDrawing.DrawStageMark(context, stage, port.Edge, edge, markEnd, markPen);
                }
            }

            if (block.Glyph == BlockGlyph.Transceiver)
            {
                Point centre = rect.Center;

                foreach ((double y, double way) in new[] { (centre.Y - 4, 1.0), (centre.Y + 4, -1.0) })
                {
                    var arrow = new StreamGeometry();

                    using (StreamGeometryContext path = arrow.Open())
                    {
                        path.BeginFigure(new Point(centre.X - (way * 4), y - 3), isFilled: false);
                        path.LineTo(new Point(centre.X + (way * 4), y));
                        path.LineTo(new Point(centre.X - (way * 4), y + 3));
                        path.EndFigure(isClosed: true);
                    }

                    context.DrawGeometry(null, thin, arrow);
                }
            }
        }

        // ###########################################################################################
        // A terminal's mark inside its block, `inset` in from the edge where it meets it - level with it on a side, under or
        // over it on the top or bottom - overlined for Q-bar. Gives how far into the block the mark reaches, for what stands
        // beside it.
        // ###########################################################################################
        private static double DrawPortMark(DrawingContext context, string mark, bool overline, Point edge, BlockEdge side, double inset, IBrush brush, Pen pen)
        {
            if (mark.Length == 0)
            {
                return inset;
            }

            FormattedText text = PinoutDrawing.TextOf(mark, brush, MarkSize, bold: false);

            Point at = side switch
            {
                BlockEdge.Left => new Point(edge.X + inset, edge.Y - (text.Height / 2)),
                BlockEdge.Right => new Point(edge.X - inset - text.Width, edge.Y - (text.Height / 2)),
                BlockEdge.Top => new Point(edge.X - (text.Width / 2), edge.Y + 1),
                _ => new Point(edge.X - (text.Width / 2), edge.Y - 1 - text.Height),
            };

            context.DrawText(text, at);

            if (overline)
            {
                double y = at.Y + (text.Height * 0.18);
                context.DrawLine(pen, new Point(at.X + 0.5, y), new Point(at.X + text.Width - 0.5, y));
            }

            return inset + text.Width;
        }

        // A 3-state output's triangle or an open-collector output's diamond, beside its mark inside the block.
        private static void DrawStageMark(DrawingContext context, LogicOutputStage stage, BlockEdge side, Point edge, double markEnd, Pen pen)
        {
            double sign = side == BlockEdge.Right ? -1 : 1;
            double x = edge.X + (sign * (markEnd + 5));
            double y = edge.Y;
            var geometry = new StreamGeometry();

            using (StreamGeometryContext path = geometry.Open())
            {
                if (stage == LogicOutputStage.ThreeState)
                {
                    path.BeginFigure(new Point(x - 3, y - 2.5), isFilled: false);
                    path.LineTo(new Point(x + 3, y - 2.5));
                    path.LineTo(new Point(x, y + 2.5));
                    path.EndFigure(isClosed: true);
                }
                else
                {
                    path.BeginFigure(new Point(x - 2.5, y - 0.5), isFilled: false);
                    path.LineTo(new Point(x, y - 3));
                    path.LineTo(new Point(x + 2.5, y - 0.5));
                    path.LineTo(new Point(x, y + 2));
                    path.EndFigure(isClosed: true);
                    path.BeginFigure(new Point(x - 2.5, y + 3), isFilled: false);
                    path.LineTo(new Point(x + 2.5, y + 3));
                    path.EndFigure(isClosed: false);
                }
            }

            context.DrawGeometry(null, pen, geometry);
        }

        // ---- switches ----------------------------------------------------------------------------------

        // A switch: its blade drawn open from the contact it turns on, both contacts over its wires' ends, and the dashed
        // link from its control's wire to the middle of the blade.
        private static void DrawSwitch(DrawingContext context, LogicSwitchShape blade, SymbolInk ink)
        {
            var outline = new Pen(ink.Line, 1.5, lineCap: PenLineCap.Round);
            var link = new Pen(ink.Line, 1, dashStyle: new DashStyle([2, 2], 0));

            context.DrawLine(link, blade.Control, blade.Link);
            context.DrawLine(outline, blade.Pivot, blade.BladeTip);

            foreach (Point contact in new[] { blade.Pivot, blade.Contact })
            {
                context.DrawEllipse(ink.Fill, new Pen(ink.Line, 1.2), contact, 2.5, 2.5);
            }
        }
    }
}
