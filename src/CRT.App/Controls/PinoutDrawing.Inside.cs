using Avalonia;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A CHIP'S INSIDES, PAINTED (owner request, 2026-11-16: "something alike the attached image, just better and prettier").
    // What is drawn is LogicChipGeometry's, the colors PinDirectionColors'; this only paints: the case as the plain one is drawn,
    // in the theme's colors (owner request, the same day: "it should have the same color as the normal drawn pinout IC"),
    // its notch and pin 1's dot; each pin a square in its direction's discreet tint (owner request, 2026-11-17 - its expected
    // reading's until then), edged in its ink, with its number on it and its name outside (the datasheet's where the board
    // names none - Show); the wires of each pin's net in its direction's wire
    // color, those inside a unit in the plain one; the symbols (PinoutDrawing.Symbols.cs); and a pin's badge where a stub is
    // joined to it. NO NAME IS WRITTEN INSIDE (owner request, the same day: "I do not think this text inside the IC is
    // needed"). The pin under the pointer, or marked by the host, is lit as on the plain case - the strip behind its square
    // and name, no ring round the square (owner request, the same day: "that is not required for the pin-box, as it has the
    // blue'ish hover") - and so is its net inside: its wires drawn wider with a glow, its badges ringed. The numbers are not
    // bold (owner request, the same day: "can we try without it being bold?").
    //
    // UNDER IT, what the colors mean - only the directions shown; not what the chip is (owner request, the same day: "the name is
    // already given elsewhere (e.g. in header line)").
    // ###########################################################################################
    public sealed partial class PinoutDrawing
    {
        private const double BadgeNumberSize = 8.5;
        private const double NumberOnSquareSize = 10;
        private const double CaptionSize = 11;
        private const double CaptionGap = 8;
        private const double LegendSwatch = 10;
        private const double LegendSwatchGap = 4;
        private const double LegendItemGap = 12;

        private const string LegendLead = "Pin direction:";

        private LogicInside? thisInside;
        private IReadOnlyDictionary<int, PinDirectionKind> thisDirections = new Dictionary<int, PinDirectionKind>();

        internal LogicInside? InsideForTests => this.thisInside;

        internal IReadOnlyDictionary<int, PinDirectionKind> DirectionsForTests => this.thisDirections;

        // The directions the legend names: those of the pins drawn, in this order.
        internal IReadOnlyList<PinDirectionKind> LegendDirectionsForTests => this.LegendDirections();

        // A chip CRT knows with its insides, any other DIP in the same case empty (LogicChipGeometry.EmptyCaseFor), each pin in
        // its direction's color (PinDirectionKinds.ByPin); a single row (a SIP, a TO-92) none - the plain case's.
        private void ShowInside(LogicChip? chip, string? package, IReadOnlyList<ComponentPinEntry> pins)
        {
            this.thisInside = chip is not null ? LogicChipGeometry.For(chip, this.thisNameWidth) : LogicChipGeometry.EmptyCaseFor(package, this.thisNameWidth);
            this.thisDirections = this.thisInside is { } inside ? PinDirectionKinds.ByPin(inside.Case.Pins.Count, pins, chip) : new Dictionary<int, PinDirectionKind>();
            this.thisGateMarks.Clear();
        }

        // ###########################################################################################
        // Each gate's marks, and each mark's inked bounds, worked out once - per Show, as the gates are made again then (code
        // review, 2026-11-17: at every render, so at every pin the pointer passed over, MarksOf fitted every mark again and each
        // digit's geometry was built again). A text's bounds depend on nothing but the text: the font and size are fixed.
        // ###########################################################################################
        private readonly Dictionary<LogicGateShape, IReadOnlyList<LogicGateMark>> thisGateMarks = new(ReferenceEqualityComparer.Instance);
        private readonly Dictionary<string, Rect> thisInkedMarks = new(StringComparer.Ordinal);

        private IReadOnlyList<LogicGateMark> MarksOf(LogicGateShape gate)
        {
            if (!this.thisGateMarks.TryGetValue(gate, out IReadOnlyList<LogicGateMark>? marks))
            {
                marks = LogicChipGeometry.MarksOf(gate);
                this.thisGateMarks[gate] = marks;
            }

            return marks;
        }

        private Rect InkedBoundsOf(FormattedText text, string mark)
        {
            if (!this.thisInkedMarks.TryGetValue(mark, out Rect inked))
            {
                inked = text.BuildGeometry(default)?.Bounds ?? new Rect(0, 0, text.Width, text.Height);
                this.thisInkedMarks[mark] = inked;
            }

            return inked;
        }

        // ###########################################################################################
        // The directions the legend names - of the pins drawn, in this order - and NONE while no pin has a direction a color
        // says (code review, 2026-11-17): a case of grey pins alone - a DIP-40 just given its package - read "Pin direction:
        // Other", which explains nothing.
        // ###########################################################################################
        private List<PinDirectionKind> LegendDirections()
        {
            List<PinDirectionKind> shown = new[] { PinDirectionKind.Input, PinDirectionKind.Output, PinDirectionKind.InputOutput, PinDirectionKind.Other }
                .Where(this.thisDirections.Values.Contains)
                .ToList();

            return shown.Any(kind => kind != PinDirectionKind.Other) ? shown : [];
        }

        // A pin's direction - and the plain one for a wire inside a unit (pin 0).
        private PinDirectionKind DirectionOf(int pin) => this.thisDirections.TryGetValue(pin, out PinDirectionKind kind) ? kind : PinDirectionKind.Other;

        // ---- the legend under the case --------------------------------------------------------------

        // Where the legend starts: under the left pins' squares.
        private static double CaptionLeftOf(PinoutLayout layout) => layout.Body.Left - LogicChipGeometry.LeadGap - LogicChipGeometry.Square;

        // The case's size and the legend's under it - worked out once per Show (code review, 2026-11-17: it measured the
        // legend's words again at every measure and every render, so at every pin the pointer passed over).
        private Size WithCaption(Size size)
        {
            if (this.thisInside is not { } inside || this.LegendDirections().Count == 0)
            {
                return size;
            }

            double height = PinoutDrawing.TextOf(LegendLead, Brushes.Black, CaptionSize, bold: false).Height;

            return new Size(
                Math.Ceiling(Math.Max(size.Width, PinoutDrawing.CaptionLeftOf(inside.Case) + this.LegendWidth())),
                Math.Ceiling(size.Height + CaptionGap + height));
        }

        private double LegendWidth() =>
            PinoutDrawing.TextOf(LegendLead, Brushes.Black, CaptionSize, bold: false).Width +
            this.LegendDirections().Sum(level =>
                LegendItemGap + LegendSwatch + LegendSwatchGap + PinoutDrawing.TextOf(PinDirectionColors.WordOf(level), Brushes.Black, CaptionSize, bold: false).Width);

        private void RenderCaption(DrawingContext context, PinoutLayout layout, IBrush foreground)
        {
            if (this.LegendDirections().Count == 0)
            {
                return;
            }

            double x = PinoutDrawing.CaptionLeftOf(layout);
            double y = layout.Body.Bottom + CaptionGap;

            FormattedText lead = PinoutDrawing.TextOf(LegendLead, foreground, CaptionSize, bold: false);
            context.DrawText(lead, new Point(x, y));
            x += lead.Width;

            foreach (PinDirectionKind level in this.LegendDirections())
            {
                x += LegendItemGap;
                var swatch = new Rect(x, y + ((lead.Height - LegendSwatch) / 2), LegendSwatch, LegendSwatch);
                context.DrawRectangle(this.FillOf(level), new Pen(this.InkOf(level), 1), new RoundedRect(swatch, 2));
                x += LegendSwatch + LegendSwatchGap;

                FormattedText word = PinoutDrawing.TextOf(PinDirectionColors.WordOf(level), foreground, CaptionSize, bold: false);
                context.DrawText(word, new Point(x, y));
                x += word.Width;
            }
        }

        // ---- the case and its insides ---------------------------------------------------------------

        private void RenderInside(DrawingContext context, LogicInside inside)
        {
            PinoutLayout layout = inside.Case;
            Rect body = layout.Body;
            IBrush foreground = this.Brush("Fg", Brushes.Black);

            // Clear, not nothing: only what is drawn takes the pointer.
            if (this.PinsCanBePicked)
            {
                context.FillRectangle(Brushes.Transparent, new Rect(this.NaturalSize));
            }

            // The pin shown and the pin under the pointer, behind everything - as on the plain case.
            if (this.thisMarkedPin is int marked && layout.Pins.FirstOrDefault(pin => pin.Number == marked) is { } markedPin)
            {
                context.DrawRectangle(this.Brush("Form_Selected_Bg", Brushes.LightBlue), null, PinoutGeometry.AreaOf(layout, markedPin), 2, 2);
            }

            if (this.thisHoveredPin is int hovered && hovered != this.thisMarkedPin && layout.Pins.FirstOrDefault(pin => pin.Number == hovered) is { } hoveredPin)
            {
                context.DrawRectangle(this.Brush("Form_Selected_Hover_Bg", Brushes.SkyBlue), null, PinoutGeometry.AreaOf(layout, hoveredPin), 2, 2);
            }

            IBrush card = this.Brush("Card_Bg", Brushes.White);

            this.RenderBody(context, body, foreground, card);
            this.RenderPins(context, layout, foreground);

            // The wires, then the lit ones over them - wider, with a glow.
            foreach (LogicWire wire in inside.Wires)
            {
                PinoutDrawing.DrawWire(context, wire, this.WireColorOf(this.DirectionOf(wire.Pin)), lit: false);
            }

            foreach (LogicWire wire in inside.Wires.Where(wire => this.IsLit(wire.Pin)))
            {
                PinoutDrawing.DrawWire(context, wire, this.WireColorOf(this.DirectionOf(wire.Pin)), lit: true);
            }

            var ink = new SymbolInk(foreground, card, this.Brush("Text_Muted_Fg", Brushes.Gray), this.WireColorOf(this.DirectionOf(0)));

            foreach (ILogicSymbol symbol in inside.Symbols)
            {
                switch (symbol)
                {
                    case LogicGateShape gate:
                        this.DrawGate(context, gate, ink);
                        break;

                    case LogicBlockShape block:
                        PinoutDrawing.DrawBlock(context, block, ink);
                        break;

                    case LogicSwitchShape blade:
                        PinoutDrawing.DrawSwitch(context, blade, ink);
                        break;
                }
            }

            foreach (LogicPinBadge badge in inside.Badges)
            {
                this.DrawBadge(context, badge, lit: this.IsLit(badge.Pin), foreground);
            }

            this.RenderCaption(context, layout, foreground);
        }

        // A wire inside a unit (pin 0) is never lit: no pin is its.
        private bool IsLit(int pin) => pin != 0 && (pin == this.thisHoveredPin || pin == this.thisMarkedPin);

        // ###########################################################################################
        // A direction's colors in the theme shown (PinDirectionColors): its ink - the edges - its square's tint, the number on it,
        // and its wires.
        // ###########################################################################################
        private IBrush InkOf(PinDirectionKind level) => this.Brush(PinDirectionColors.InkKeyOf(level), Brushes.Gray);

        private IBrush FillOf(PinDirectionKind level) => this.Brush(PinDirectionColors.FillKeyOf(level), Brushes.WhiteSmoke);

        private IBrush NumberOf(PinDirectionKind level) => this.Brush(PinDirectionColors.NumberKeyOf(level), Brushes.Black);

        private Color WireColorOf(PinDirectionKind level) =>
            this.Brush(PinDirectionColors.WireKeyOf(level), Brushes.Gray) is ISolidColorBrush solid ? solid.Color : Colors.Gray;

        // ###########################################################################################
        // The case as the plain one is drawn (owner request, 2026-11-16: "it should have the same color as the normal drawn
        // pinout IC"): the theme's card ground with its text color for the edge, the notch cut into its top, the dot by pin 1
        // - in its top left corner, clear of the wires.
        // ###########################################################################################
        private void RenderBody(DrawingContext context, Rect body, IBrush foreground, IBrush card)
        {
            var outline = new Pen(foreground, 1);
            context.DrawRectangle(card, outline, body, 3, 3);

            double notch = LogicChipGeometry.NotchRadius;
            double dot = LogicChipGeometry.FirstPinDot;

            using (context.PushClip(body))
            {
                context.DrawGeometry(this.Brush("Bg", Brushes.White), outline, new EllipseGeometry(new Rect(body.Center.X - notch, body.Top - notch, 2 * notch, 2 * notch)));
            }

            context.DrawEllipse(this.Brush("Schematics_FirstPin", Brushes.IndianRed), null, new Point(body.Left + dot, body.Top + dot), LogicChipGeometry.DotRadius, LogicChipGeometry.DotRadius);
        }

        // Each pin: its lead from the case, its square in its direction's tint with its number, its name outside.
        private void RenderPins(DrawingContext context, PinoutLayout layout, IBrush foreground)
        {
            Rect body = layout.Body;

            foreach (PinoutPin pin in layout.Pins)
            {
                PinDirectionKind level = this.DirectionOf(pin.Number);
                Rect square = pin.Lead;

                context.DrawRectangle(
                    foreground,
                    null,
                    pin.Side == PinSide.Left
                        ? new Rect(square.Right, square.Center.Y - 3, body.Left - square.Right, 6)
                        : new Rect(body.Right, square.Center.Y - 3, square.Left - body.Right, 6));

                context.DrawRectangle(this.FillOf(level), new Pen(this.InkOf(level), 1), new RoundedRect(square, 3));

                FormattedText number = PinoutDrawing.TextOf(
                    pin.Number.ToString(CultureInfo.InvariantCulture), this.NumberOf(level), NumberOnSquareSize, bold: false);
                context.DrawText(number, new Point(square.Center.X - (number.Width / 2), square.Center.Y - (number.Height / 2)));

                if (!this.thisNames.TryGetValue(pin.Number, out string? name) || name.Length == 0)
                {
                    continue;
                }

                FormattedText nameText = PinoutDrawing.Text(name, foreground, this.thisNameWidth - PinoutGeometry.NameGaps, bold: false);

                context.DrawText(
                    nameText,
                    pin.Side == PinSide.Left
                        ? new Point(pin.NameAnchor.X - nameText.Width, pin.NameAnchor.Y - (nameText.Height / 2))
                        : new Point(pin.NameAnchor.X, pin.NameAnchor.Y - (nameText.Height / 2)));
            }
        }

        private static void DrawWire(DrawingContext context, LogicWire wire, Color line, bool lit)
        {
            var geometry = new StreamGeometry();

            using (StreamGeometryContext path = geometry.Open())
            {
                path.BeginFigure(wire.Points[0], isFilled: false);

                foreach (Point point in wire.Points.Skip(1))
                {
                    path.LineTo(point);
                }

                path.EndFigure(isClosed: false);
            }

            if (lit)
            {
                context.DrawGeometry(null, new Pen(new SolidColorBrush(Color.FromArgb(0x50, line.R, line.G, line.B)), 7, lineCap: PenLineCap.Round, lineJoin: PenLineJoin.Round), geometry);
            }

            context.DrawGeometry(null, new Pen(new SolidColorBrush(line), lit ? 2.6 : WireWidth, lineCap: PenLineCap.Flat, lineJoin: PenLineJoin.Round), geometry);
        }

        // How thick a wire is drawn - a stub to a tag outside a gate's back too.
        private const double WireWidth = 1.6;

        // A pin's badge at a symbol: its number on its direction's tint, edged in its ink, as its square is - ringed while the pin is lit.
        private void DrawBadge(DrawingContext context, LogicPinBadge badge, bool lit, IBrush ring)
        {
            PinDirectionKind level = this.DirectionOf(badge.Pin);

            context.DrawRectangle(this.FillOf(level), new Pen(this.InkOf(level), 1), new RoundedRect(badge.Area, 3));

            FormattedText number = PinoutDrawing.TextOf(badge.Pin.ToString(CultureInfo.InvariantCulture), this.NumberOf(level), BadgeNumberSize, bold: false);
            context.DrawText(number, new Point(badge.Area.Center.X - (number.Width / 2), badge.Area.Center.Y - (number.Height / 2)));

            if (lit)
            {
                context.DrawRectangle(null, new Pen(ring, 1.5), new RoundedRect(badge.Area.Inflate(2), 4));
            }
        }

        private static FormattedText TextOf(string text, IBrush brush, double size, bool bold) =>
            new(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(FontFamily.Default, weight: bold ? FontWeight.Bold : FontWeight.Normal), size, brush);
    }
}
