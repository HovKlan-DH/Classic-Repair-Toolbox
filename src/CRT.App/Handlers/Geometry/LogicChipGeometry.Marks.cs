using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // What a mark inside a gate is: a comparator's "+" or "-", a multiplexer's input number, or what the gate's output
    // stage or inputs are - 3-state, open collector, Schmitt trigger.
    public enum GateMarkKind
    {
        Plus,
        Minus,
        Text,
        ThreeState,
        OpenCollector,
        Schmitt,
    }

    // ###########################################################################################
    // A mark drawn inside a gate: what it is, the area it is drawn to fill (a text's ink centred in it), the points its ink
    // reaches out to - the area's corners, or a sign's arm tips - and its text, for a number.
    // ###########################################################################################
    public sealed record LogicGateMark(GateMarkKind Kind, Rect Area, IReadOnlyList<Point> Ink, string Text = "");

    // ###########################################################################################
    // *** EVERY MARK INSIDE A GATE KEEPS CLEAR OF ITS OUTLINE (owner request, 2026-11-17: "the +/- is too close to its
    // borders ... we need a high quality with a lot of details. There could be others with same issue"). *** A 556's
    // comparators had their "-" on the slanted edge, the 257's multiplexers their "0" a pixel under the top edge, the 7414's
    // Schmitt mark and the 7406's open-collector mark ran into their inverters' sides. So where a mark goes is decided here,
    // not by the painter: each in an area of its own (MarksOf), its ink MarkClear from the outline's ink - the strokes'
    // widths, GateOutline and MarkStroke, counted - and a gate is made as large as its marks need: a comparator as tall as
    // its signs take (ComparatorHalfHeight), a multiplexer as its outer numbers take (MuxHeightFor), a buffer BufferHeight
    // by BodyWidthOf(Buffer) with its stage or Schmitt mark drawn as large as fits there, down to half. A gate too small for
    // a mark even at half its size - the 240's and 244's buffers, a row's half apart - is drawn without it, its enable
    // saying it is 3-state. PinoutDrawing paints each mark to fill its area; LogicChipGeometryTests hold every gate's marks
    // to its outline. The curved gates (AND, OR, XOR) have their marks in their middle, with room to spare.
    // ###########################################################################################
    public static partial class LogicChipGeometry
    {
        public const double MarkClear = 2.75;
        public const double MarkStroke = 1;

        // From a gate's back to where a mark at an input begins: the clearance and the two strokes' halves.
        public const double MarkInset = LogicChipGeometry.MarkClear + (LogicChipGeometry.GateOutline / 2) + (LogicChipGeometry.MarkStroke / 2);

        // A buffer's size, and the room a multiplexer's number takes - a digit at the marks' size (6.1 high, 4.7 wide in Segoe
        // UI), with a little to spare.
        public const double BufferHeight = 24;
        public const double DigitWidth = 5;
        public const double DigitHeight = 6.5;

        // The stage and Schmitt marks at their full size; how small one may be drawn to fit.
        private const double ThreeStateWidth = 7;
        private const double ThreeStateHeight = 6;
        private const double OpenCollectorSize = 3.4;
        private const double SchmittWidth = 10;
        private const double SchmittHeight = 6;
        private const double LeastMarkScale = 0.5;

        // How much room a mark's ink must have from the outline's centre line.
        private static double RoomNeeded => LogicChipGeometry.MarkClear + (LogicChipGeometry.GateOutline / 2) + (LogicChipGeometry.MarkStroke / 2);

        // ###########################################################################################
        // The outline of a gate with straight edges - a buffer's or comparator's triangle, a multiplexer's trapezoid - as if its
        // back stood at x 0 and it pointed right, its middle at y 0: `width` long, `half` up and down at its back. Null for a
        // curved one.
        // ###########################################################################################
        private static Point[]? StraightOutline(GateSymbol shape, double width, double half)
        {
            double taper = LogicChipGeometry.MuxTaper(2 * half);

            return shape switch
            {
                GateSymbol.Buffer or GateSymbol.Comparator => [new(0, -half), new(width, 0), new(0, half)],
                GateSymbol.Mux => [new(0, -half), new(width, -half + taper), new(width, half - taper), new(0, half)],
                _ => null,
            };
        }

        // The least distance from any of `ink` to the outline's edges, each point inside it - negative for one outside.
        private static double RoomAround(Point[] outline, IEnumerable<Point> ink) =>
            ink.Min(point =>
            {
                double least = double.PositiveInfinity;

                for (int index = 0; index < outline.Length; index++)
                {
                    Point a = outline[index];
                    Point b = outline[(index + 1) % outline.Length];
                    double length = Math.Sqrt(((b.X - a.X) * (b.X - a.X)) + ((b.Y - a.Y) * (b.Y - a.Y)));

                    // The outline runs clockwise on screen (y down), so the inside is on each edge's right: positive.
                    double side = (((b.X - a.X) * (point.Y - a.Y)) - ((b.Y - a.Y) * (point.X - a.X))) / length;
                    least = Math.Min(least, side);
                }

                return least;
            });

        private static bool Fits(Point[] outline, IEnumerable<Point> ink) => LogicChipGeometry.RoomAround(outline, ink) >= LogicChipGeometry.RoomNeeded - 1e-9;

        private static IEnumerable<Point> CornersOf(Rect area) => [area.TopLeft, area.TopRight, area.BottomRight, area.BottomLeft];

        // A sign's ink: its arm tips about `centre` - a "-" has only the two across.
        private static Point[] SignInk(Point centre, bool plus) =>
            plus
                ? [new(centre.X - LogicChipGeometry.SignArm, centre.Y), new(centre.X + LogicChipGeometry.SignArm, centre.Y), new(centre.X, centre.Y - LogicChipGeometry.SignArm), new(centre.X, centre.Y + LogicChipGeometry.SignArm)]
                : [new(centre.X - LogicChipGeometry.SignArm, centre.Y), new(centre.X + LogicChipGeometry.SignArm, centre.Y)];

        // A multiplexer's number at an input `y` from the middle, as if the back stood at x 0.
        private static Rect DigitArea(double y) =>
            new(LogicChipGeometry.MarkInset, y - (LogicChipGeometry.DigitHeight / 2), LogicChipGeometry.DigitWidth, LogicChipGeometry.DigitHeight);

        // ###########################################################################################
        // How far a comparator reaches up and down from its middle, its inputs `apart` from it: the least, in half pixels,
        // that leaves each sign's ink MarkClear from the outline's.
        // ###########################################################################################
        public static double ComparatorHalfHeight(double apart)
        {
            double width = LogicChipGeometry.BodyWidthOf(GateSymbol.Comparator);
            Point[] ink = [.. LogicChipGeometry.SignInk(new Point(LogicChipGeometry.SignInset, -apart), plus: true), .. LogicChipGeometry.SignInk(new Point(LogicChipGeometry.SignInset, apart), plus: true)];

            for (double half = apart + LogicChipGeometry.SignArm; half < apart + 100; half += 0.5)
            {
                if (LogicChipGeometry.Fits(LogicChipGeometry.StraightOutline(GateSymbol.Comparator, width, half)!, ink))
                {
                    return half;
                }
            }

            throw new InvalidOperationException("No comparator is tall enough for its signs.");
        }

        // ###########################################################################################
        // How tall a multiplexer `width` long is, its outer inputs `span` apart: as the datasheets draw one, 16 more than that,
        // or as much more, in whole pixels, as its top and bottom numbers need to keep clear of the sloping edges.
        // ###########################################################################################
        public static double MuxHeightFor(double span, double width)
        {
            IEnumerable<Point> ink = LogicChipGeometry.CornersOf(LogicChipGeometry.DigitArea(-span / 2))
                .Concat(LogicChipGeometry.CornersOf(LogicChipGeometry.DigitArea(span / 2)));

            for (double height = span + 16; height < span + 100; height++)
            {
                if (LogicChipGeometry.Fits(LogicChipGeometry.StraightOutline(GateSymbol.Mux, width, height / 2)!, ink))
                {
                    return height;
                }
            }

            throw new InvalidOperationException("No multiplexer is tall enough for its numbers.");
        }

        // ###########################################################################################
        // Every mark inside `gate`, where it is drawn: a comparator's signs and a multiplexer's numbers at their inputs, and
        // its stage or Schmitt mark - against a triangle's back, by a multiplexer's front, in the middle of a curved gate -
        // as large as fits with its room, down to half its size; none at all when even that does not fit.
        // ###########################################################################################
        public static IReadOnlyList<LogicGateMark> MarksOf(LogicGateShape gate)
        {
            ArgumentNullException.ThrowIfNull(gate);

            double facing = gate.Facing == GateFacing.Right ? 1 : -1;
            double half = gate.BodyHeight / 2;
            double width = gate.BodyWidth;
            Point[]? outline = LogicChipGeometry.StraightOutline(gate.Shape, width, half);
            var marks = new List<LogicGateMark>();

            // From the gate's own coordinates - back at x 0, pointing right, middle at y 0 - to the drawing's.
            Point Out(Point local) => new(gate.BackX + (facing * local.X), gate.CentreY + local.Y);

            Rect AreaOut(Rect local)
            {
                Point one = Out(local.TopLeft);
                Point other = Out(local.BottomRight);
                return new Rect(Math.Min(one.X, other.X), Math.Min(one.Y, other.Y), Math.Abs(other.X - one.X), Math.Abs(other.Y - one.Y));
            }

            LogicGateMark Placed(GateMarkKind kind, Rect local, IEnumerable<Point> ink, string text = "") =>
                new(kind, AreaOut(local), [.. ink.Select(Out)], text);

            for (int index = 0; index < gate.Inputs.Count; index++)
            {
                string mark = index < gate.InputMarks.Count ? gate.InputMarks[index] : string.Empty;
                double y = gate.Inputs[index].Y - gate.CentreY;

                if (gate.Shape == GateSymbol.Comparator && mark is "+" or "−")
                {
                    var centre = new Point(LogicChipGeometry.SignInset, y);
                    var area = new Rect(centre.X - LogicChipGeometry.SignArm, y - LogicChipGeometry.SignArm, 2 * LogicChipGeometry.SignArm, 2 * LogicChipGeometry.SignArm);
                    marks.Add(Placed(mark == "+" ? GateMarkKind.Plus : GateMarkKind.Minus, area, LogicChipGeometry.SignInk(centre, mark == "+")));
                }
                else if (mark.Length > 0)
                {
                    Rect area = LogicChipGeometry.DigitArea(y);
                    marks.Add(Placed(GateMarkKind.Text, area, LogicChipGeometry.CornersOf(area), mark));
                }
            }

            (GateMarkKind Kind, double Width, double Height)? inner =
                gate.Stage == LogicOutputStage.ThreeState ? (GateMarkKind.ThreeState, LogicChipGeometry.ThreeStateWidth, LogicChipGeometry.ThreeStateHeight)
                : gate.Stage == LogicOutputStage.OpenCollector ? (GateMarkKind.OpenCollector, 2 * LogicChipGeometry.OpenCollectorSize, (2 * LogicChipGeometry.OpenCollectorSize) + 1.5)
                : gate.Schmitt ? (GateMarkKind.Schmitt, LogicChipGeometry.SchmittWidth, LogicChipGeometry.SchmittHeight)
                : null;

            if (inner is not { } wanted)
            {
                return marks;
            }

            for (double scale = 1; scale >= LogicChipGeometry.LeastMarkScale - 1e-9; scale -= 0.05)
            {
                double w = wanted.Width * scale;
                double h = wanted.Height * scale;
                double xorShift = gate.Shape == GateSymbol.Xor ? 5 : 0;

                // Where its middle stands along the gate.
                double middle = gate.Shape switch
                {
                    GateSymbol.Buffer or GateSymbol.Comparator => LogicChipGeometry.MarkInset + (w / 2),
                    GateSymbol.Mux => width - LogicChipGeometry.MarkInset - (w / 2),
                    GateSymbol.Xor => xorShift + ((width - xorShift) * 0.48),
                    _ => width * 0.45,
                };

                var area = new Rect(middle - (w / 2), -h / 2, w, h);

                if (outline is null || LogicChipGeometry.Fits(outline, LogicChipGeometry.CornersOf(area)))
                {
                    marks.Add(Placed(wanted.Kind, area, LogicChipGeometry.CornersOf(area)));
                    break;
                }
            }

            return marks;
        }
    }
}
