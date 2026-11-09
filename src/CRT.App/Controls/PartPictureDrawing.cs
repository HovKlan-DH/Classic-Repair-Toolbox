using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A LEADED PART DRAWN AS IT LOOKS, EACH LEAD LABELLED - the pinout picture of a capacitor, a diode, a transistor, a
    // regulator, a crystal, a bead or coil, an EMI filter, a resistor pack or a pin header (owner request, 2026-11-09: "could
    // you also draw an axial and radial electrolytic capacitor in same style ... and a disc capacitor as a pinout image and
    // mark (if applicable) its plus and minus side? And how about transistors? And diodes?"). In the style of the
    // resistor's picture (ResistorBandsDrawing): a part lying on the page - a soft shadow, round tinned leads, the body
    // shaded as the cylinder, disc or block it is - in its own colors on either theme; the labels under the leads in the
    // theme's text color.
    //
    // What is drawn is PartPicture's, every position PartPictureGeometry's; this only paints. The component popup and the
    // Contribute tab's "Pinout" page show it in the place of PinoutDrawing, so it does what that does: with PinsCanBePicked
    // a lead lights up under the pointer and a click raises PinPicked; the host lights one too (HoverPin, MarkPin), and the
    // pointer moving onto or off a lead raises PinHovered (owner request, 2026-11-09: the list's row lit too). It is
    // drawn smaller, never cut, where its place is narrower than it.
    // ###########################################################################################
    public sealed class PartPictureDrawing : Control
    {
        public static readonly RoutedEvent<PinPickedEventArgs> PinPickedEvent =
            RoutedEvent.Register<PartPictureDrawing, PinPickedEventArgs>(nameof(PinPicked), RoutingStrategies.Bubble);

        public static readonly RoutedEvent<PinHoveredEventArgs> PinHoveredEvent =
            RoutedEvent.Register<PartPictureDrawing, PinHoveredEventArgs>(nameof(PinHovered), RoutingStrategies.Bubble);

        private const double NumberSize = 11;
        private const double NameSize = 10.5;
        private const double SignSize = 15;
        private const double CaptionSize = 10.5;
        private const double CaptionGap = 4;

        private PartPicture.Picture? thisPicture;
        private PartPictureGeometry.Layout? thisLayout;
        private FormattedText? thisCaption;

        private int? thisHoveredPin;
        private int? thisMarkedPin;

        // How much smaller than its own size it is drawn, for the place it has.
        private double thisScale = 1;

        public event EventHandler<PinPickedEventArgs>? PinPicked
        {
            add => this.AddHandler(PartPictureDrawing.PinPickedEvent, value);
            remove => this.RemoveHandler(PartPictureDrawing.PinPickedEvent, value);
        }

        public event EventHandler<PinHoveredEventArgs>? PinHovered
        {
            add => this.AddHandler(PartPictureDrawing.PinHoveredEvent, value);
            remove => this.RemoveHandler(PartPictureDrawing.PinHoveredEvent, value);
        }

        public bool PinsCanBePicked { get; set; }

        internal PartPicture.Picture? PictureForTests => this.thisPicture;

        internal PartPictureGeometry.Layout? LayoutForTests => this.thisLayout;

        internal int? LitPinForTests => this.thisHoveredPin ?? this.thisMarkedPin;

        // The picture drawn - null draws nothing and takes no room.
        public void Show(PartPicture.Picture? picture)
        {
            this.thisPicture = picture;
            this.thisLayout = picture is null ? null : PartPictureGeometry.For(picture);
            this.thisCaption = picture is { Caption.Length: > 0 }
                ? PartPictureDrawing.Text(picture.Caption, Brushes.Black, CaptionSize, FontWeight.Normal)
                : null;
            this.thisHoveredPin = null;
            this.thisMarkedPin = null;
            this.IsVisible = this.thisLayout is not null;
            this.InvalidateMeasure();
            this.InvalidateVisual();
        }

        // The pin of what the host shows - null for none. It stays lit until the host says otherwise.
        public void MarkPin(int? number)
        {
            if (this.thisMarkedPin != number)
            {
                this.thisMarkedPin = number;
                this.InvalidateVisual();
            }
        }

        // The pin under the pointer in the host's list - null for none.
        public void HoverPin(int? number)
        {
            if (this.thisHoveredPin != number)
            {
                this.thisHoveredPin = number;
                this.InvalidateVisual();
            }
        }

        // Its own size: the picture, what its leads' letters stand for under it, and the caption under that.
        internal Size NaturalSize
        {
            get
            {
                if (this.thisLayout is not { } layout || this.thisPicture is not { } picture)
                {
                    return default;
                }

                double width = layout.Size.Width;
                double height = layout.Size.Height;

                if (picture.LegendLines.Count > 0)
                {
                    width = Math.Max(width, PartPictureDrawing.LegendWidth(picture));
                    height += CaptionGap + (picture.LegendLines.Count * LegendLineHeight);
                }

                if (this.thisCaption is { } caption)
                {
                    width = Math.Max(width, Math.Ceiling(caption.Width));
                    height += CaptionGap + Math.Ceiling(caption.Height);
                }

                return new Size(width, height);
            }
        }

        // A line of the legend: "E = Emitter", its letter bold.
        private const double LegendLineHeight = 15;
        private const double LegendSize = 11;

        private static double LegendWidth(PartPicture.Picture picture) =>
            picture.LegendLines.Max(line =>
                Math.Ceiling(PartPictureDrawing.Text(line.Letter, Brushes.Black, LegendSize, FontWeight.Bold).Width +
                             PartPictureDrawing.Text(" = " + line.Meaning, Brushes.Black, LegendSize, FontWeight.Normal).Width));

        protected override Size MeasureOverride(Size availableSize)
        {
            Size natural = this.NaturalSize;

            this.thisScale = natural.Width > 0 && availableSize.Width < natural.Width ? Math.Max(0.4, availableSize.Width / natural.Width) : 1;

            return new Size(natural.Width * this.thisScale, natural.Height * this.thisScale);
        }

        // ---- the pointer ------------------------------------------------------------------------

        private PartPictureGeometry.LeadPlace? LeadUnder(PointerEventArgs e) =>
            this.thisLayout is { } layout ? PartPictureGeometry.LeadAt(layout, e.GetPosition(this) / this.thisScale) : null;

        protected override void OnPointerMoved(PointerEventArgs e)
        {
            base.OnPointerMoved(e);

            if (this.PinsCanBePicked)
            {
                this.HoverOver(this.LeadUnder(e)?.Number);
            }
        }

        protected override void OnPointerExited(PointerEventArgs e)
        {
            base.OnPointerExited(e);
            this.HoverOver(null);
        }

        protected override void OnPointerPressed(PointerPressedEventArgs e)
        {
            base.OnPointerPressed(e);

            if (this.PinsCanBePicked && e.GetCurrentPoint(this).Properties.IsLeftButtonPressed && this.LeadUnder(e) is { } lead)
            {
                this.RaiseEvent(new PinPickedEventArgs(PartPictureDrawing.PinPickedEvent, lead.Number));
                e.Handled = true;
            }
        }

        private void HoverOver(int? number)
        {
            if (this.thisHoveredPin != number)
            {
                this.thisHoveredPin = number;
                this.Cursor = number is null ? null : new Cursor(StandardCursorType.Hand);
                this.InvalidateVisual();
                this.RaiseEvent(new PinHoveredEventArgs(PartPictureDrawing.PinHoveredEvent, number));
            }
        }

        // ---- drawing ----------------------------------------------------------------------------

        public override void Render(DrawingContext context)
        {
            if (this.thisLayout is not { } layout || this.thisPicture is not { } picture)
            {
                return;
            }

            using (context.PushTransform(Matrix.CreateScale(this.thisScale, this.thisScale)))
            {
                // Clear, not nothing: only what is drawn takes the pointer.
                if (this.PinsCanBePicked)
                {
                    context.FillRectangle(Brushes.Transparent, new Rect(this.NaturalSize));
                }

                this.DrawLit(context, layout);

                switch (picture.Shape)
                {
                    case PartPicture.Shape.RadialElectrolytic:
                        PartPictureDrawing.DrawRadialElectrolytic(context, layout, picture);
                        break;
                    case PartPicture.Shape.AxialElectrolytic:
                        PartPictureDrawing.DrawAxialElectrolytic(context, layout, picture);
                        break;
                    case PartPicture.Shape.DiscCapacitor:
                        PartPictureDrawing.DrawDisc(context, layout, picture);
                        break;
                    case PartPicture.Shape.RadialCapacitor:
                        PartPictureDrawing.DrawDipped(context, layout, picture, "#D9A441", "#3A2A08");
                        break;
                    case PartPicture.Shape.EmiFilter:
                        PartPictureDrawing.DrawDipped(context, layout, picture, "#7A5236", "#F3E2C8");
                        break;
                    case PartPicture.Shape.AxialCapacitor:
                        PartPictureDrawing.DrawAxialBody(context, layout, picture, "#E3CB72", radiusShare: 0.42, markingColor: "#3A2E10");
                        break;
                    case PartPicture.Shape.GlassDiode:
                        PartPictureDrawing.DrawGlassDiode(context, layout);
                        break;
                    case PartPicture.Shape.EpoxyDiode:
                        PartPictureDrawing.DrawAxialBody(context, layout, picture, "#262626", radiusShare: 0.3, markingColor: null, bandColor: "#CFCFCF");
                        break;
                    case PartPicture.Shape.FerriteBead:
                        PartPictureDrawing.DrawAxialBody(context, layout, picture, "#55555B", radiusShare: 0.16, markingColor: null);
                        break;
                    case PartPicture.Shape.Coil:
                        PartPictureDrawing.DrawCoil(context, layout);
                        break;
                    case PartPicture.Shape.TO92:
                        PartPictureDrawing.DrawTO92(context, layout, picture);
                        break;
                    case PartPicture.Shape.TO220:
                        PartPictureDrawing.DrawTO220(context, layout, picture);
                        break;
                    case PartPicture.Shape.Crystal:
                        PartPictureDrawing.DrawCrystal(context, layout, picture);
                        break;
                    case PartPicture.Shape.ResistorNetwork:
                        PartPictureDrawing.DrawResistorNetwork(context, layout, picture);
                        break;
                    case PartPicture.Shape.PinHeader:
                        PartPictureDrawing.DrawPinHeader(context, layout);
                        break;
                }

                this.DrawLabels(context, layout, picture);
            }
        }

        // The lead shown and the lead under the pointer, each in its own - as PinoutDrawing lights a pin.
        private void DrawLit(DrawingContext context, PartPictureGeometry.Layout layout)
        {
            if (this.thisMarkedPin is int marked && layout.Leads.FirstOrDefault(lead => lead.Number == marked) is { } markedLead)
            {
                context.DrawRectangle(this.Brush("Form_Selected_Bg", Brushes.LightBlue), null, markedLead.Area, 3, 3);
            }

            if (this.thisHoveredPin is int lit && lit != this.thisMarkedPin && layout.Leads.FirstOrDefault(lead => lead.Number == lit) is { } litLead)
            {
                context.DrawRectangle(this.Brush("Form_Selected_Hover_Bg", Brushes.SkyBlue), null, litLead.Area, 3, 3);
            }
        }

        // Under each lead its pin number, and its name under that; the caption under the picture.
        private void DrawLabels(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture)
        {
            IBrush foreground = this.Brush("Fg", Brushes.Black);

            for (int i = 0; i < layout.Leads.Count; i++)
            {
                Rect label = layout.Leads[i].Label;
                Rect nameBox = layout.Leads[i].NameBox;
                string name = picture.Leads[i].Name;

                FormattedText number = PartPictureDrawing.Text(layout.Leads[i].Number.ToString(CultureInfo.InvariantCulture), foreground, NumberSize, FontWeight.Bold, label.Width);
                context.DrawText(number, new Point(label.Center.X - (number.Width / 2), label.Top));

                if (name.Length > 0)
                {
                    // A lone "+" or "-" is a sign to be seen at a glance, not a word: larger, and bold.
                    bool sign = name is "+" or "-";
                    FormattedText text = PartPictureDrawing.Text(name, foreground, sign ? SignSize : NameSize, sign ? FontWeight.Bold : FontWeight.Normal, nameBox.Width);
                    context.DrawText(text, new Point(nameBox.Center.X - (text.Width / 2), nameBox.Top - (sign ? 3 : 0)));
                }
            }

            // What the letters under the leads stand for (owner request, 2026-11-09: "E = Emitter, C = Collector, B = Base").
            double y = layout.Size.Height;

            if (picture.LegendLines.Count > 0)
            {
                y += CaptionGap;

                foreach (PartPicture.LegendLine line in picture.LegendLines)
                {
                    FormattedText letter = PartPictureDrawing.Text(line.Letter, foreground, LegendSize, FontWeight.Bold);
                    FormattedText meaning = PartPictureDrawing.Text(" = " + line.Meaning, foreground, LegendSize, FontWeight.Normal);
                    context.DrawText(letter, new Point(0, y));
                    context.DrawText(meaning, new Point(letter.Width, y));
                    y += LegendLineHeight;
                }
            }

            if (this.thisCaption is not null)
            {
                FormattedText text = PartPictureDrawing.Text(picture.Caption, foreground, CaptionSize, FontWeight.Normal);
                using (context.PushOpacity(0.75))
                {
                    context.DrawText(text, new Point(0, y + CaptionGap));
                }
            }
        }

        internal IReadOnlyList<string> LegendShownForTests =>
            this.thisPicture?.LegendLines.Select(line => $"{line.Letter} = {line.Meaning}").ToList() ?? [];

        // ###########################################################################################
        // *** EACH PART, AS IT LOOKS. *** Lit from above and a little from the left, as the resistor is: a soft shadow
        // under each body, its colors the real part's, what it has printed on it in the color it is printed in.
        // ###########################################################################################

        // A radial electrolytic: the can in its blue sleeve, the rubber seal under it, the light minus stripe down the
        // minus leg's side with its "-" marks, the crimp near the foot, the value printed beside the stripe.
        private static void DrawRadialElectrolytic(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture)
        {
            Rect body = layout.Body;
            var sleeve = new Rect(body.X, body.Y, body.Width, body.Height * 0.94);
            var seal = new Rect(body.X + (body.Width * 0.12), sleeve.Bottom - 3, body.Width * 0.76, body.Bottom - sleeve.Bottom + 3);
            var outline = new RoundedRect(sleeve, new CornerRadius(body.Width * 0.12, body.Width * 0.12, 5, 5));

            PartPictureDrawing.DrawLegs(context, layout);
            PartPictureDrawing.Shadow(context, outline, body.Width * 0.12);

            context.DrawRectangle(PartPictureDrawing.Solid("#2E2E30"), null, new RoundedRect(seal, 2));
            context.DrawRectangle(PartPictureDrawing.Upright(), null, new RoundedRect(seal, 2));

            using (context.PushClip(outline))
            {
                context.DrawRectangle(PartPictureDrawing.Solid("#1F4D9E"), null, sleeve);

                if (layout.Band is Rect stripe)
                {
                    context.DrawRectangle(PartPictureDrawing.Solid("#D5DAE2"), null, stripe);
                    PartPictureDrawing.MinusMarks(context, stripe, "#1F4D9E");
                }

                // The crimp round the can near its foot, and the sleeve's fold over the top.
                double crimp = body.Bottom - (body.Height * 0.17);
                context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0x70, 0, 0, 0)), null, new Rect(sleeve.Left, crimp, sleeve.Width, 2));
                context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0x45, 0xFF, 0xFF, 0xFF)), null, new Rect(sleeve.Left, crimp + 2, sleeve.Width, 1.5));
                context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0x30, 0xFF, 0xFF, 0xFF)), null, new Rect(sleeve.Left, sleeve.Top, sleeve.Width, 5));

                PartPictureDrawing.Print(context, picture.Marking, layout.Marking, "#F4F6FA", 11.5);

                context.DrawRectangle(PartPictureDrawing.Upright(), null, sleeve);
            }

            context.DrawRectangle(null, PartPictureDrawing.Edge(), outline);
        }

        // ###########################################################################################
        // An axial electrolytic: the can lying across in its blue sleeve, the bare aluminium of its bottom at the minus
        // end, the rubber seal standing out of its plus end - and by that end THE GROOVE pressed round the can, the
        // outline pinched in and its two walls lit and shaded (owner request, 2026-11-09: no minus band, "an indention
        // ring at the plus-side"). The value printed in the middle.
        // ###########################################################################################
        private static void DrawAxialElectrolytic(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture)
        {
            Rect body = layout.Body;
            bool minusLeft = picture.MarkedLead == 0;
            Rect groove = layout.Band ?? default;
            double depth = body.Height * 0.15;
            Geometry outline = PartPictureDrawing.GroovedCan(body, groove, depth, body.Height * 0.2);

            PartPictureDrawing.DrawLeads(context, layout);
            PartPictureDrawing.Shadow(context, new RoundedRect(body, body.Height * 0.2), body.Height * 0.3);

            if (layout.Tab is Rect seal)
            {
                context.DrawRectangle(PartPictureDrawing.Solid("#2E2E30"), null, new RoundedRect(seal, 2));
                context.DrawRectangle(PartPictureDrawing.Lying(), null, new RoundedRect(seal, 2));
            }

            using (context.PushGeometryClip(outline))
            {
                context.DrawRectangle(PartPictureDrawing.Solid("#1F4D9E"), null, body);

                // The bare aluminium of the can's bottom at the minus end, the sleeve's edge rolled over the plus end.
                double rim = 3;
                context.DrawRectangle(PartPictureDrawing.Solid("#C9CDD3"), null, new Rect(minusLeft ? body.Left : body.Right - rim, body.Top, rim, body.Height));
                context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0x55, 0, 0, 0)), null, new Rect(minusLeft ? body.Right - 4 : body.Left, body.Top, 4, body.Height));

                PartPictureDrawing.Print(context, picture.Marking.Replace('\n', ' '), layout.Marking, "#F4F6FA", 12);

                // The groove: its wall facing the light bright, the other in shadow, darkest at its floor.
                if (groove.Width > 0)
                {
                    context.DrawRectangle(
                        new LinearGradientBrush
                        {
                            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                            EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative),
                            GradientStops =
                            {
                                new GradientStop(Color.FromArgb(0x00, 0, 0, 0), 0),
                                new GradientStop(Color.FromArgb(0x90, 0, 0, 0), 0.42),
                                new GradientStop(Color.FromArgb(0xB0, 0, 0, 0), 0.5),
                                new GradientStop(Color.FromArgb(0x70, 0xFF, 0xFF, 0xFF), 0.62),
                                new GradientStop(Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF), 1),
                            },
                        },
                        null,
                        groove);
                }

                context.DrawRectangle(PartPictureDrawing.Lying(), null, body);
            }

            context.DrawGeometry(null, PartPictureDrawing.Edge(), outline);
        }

        // ###########################################################################################
        // A can lying across with a groove pressed round it: a rounded rectangle whose top and bottom edges dip in by
        // `depth` over the groove's width, smoothly, so the silhouette shows the ring as the real can does.
        // ###########################################################################################
        private static StreamGeometry GroovedCan(Rect body, Rect groove, double depth, double radius)
        {
            var geometry = new StreamGeometry();
            double g0 = groove.Left, g1 = groove.Right, gc = groove.Center.X, w = groove.Width;
            bool grooved = w > 0;

            using StreamGeometryContext path = geometry.Open();
            path.BeginFigure(new Point(body.Left + radius, body.Top), isFilled: true);

            if (grooved)
            {
                path.LineTo(new Point(g0, body.Top));
                path.CubicBezierTo(new Point(g0 + (w * 0.25), body.Top), new Point(gc - (w * 0.25), body.Top + depth), new Point(gc, body.Top + depth));
                path.CubicBezierTo(new Point(gc + (w * 0.25), body.Top + depth), new Point(g1 - (w * 0.25), body.Top), new Point(g1, body.Top));
            }

            path.LineTo(new Point(body.Right - radius, body.Top));
            path.ArcTo(new Point(body.Right, body.Top + radius), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
            path.LineTo(new Point(body.Right, body.Bottom - radius));
            path.ArcTo(new Point(body.Right - radius, body.Bottom), new Size(radius, radius), 0, false, SweepDirection.Clockwise);

            if (grooved)
            {
                path.LineTo(new Point(g1, body.Bottom));
                path.CubicBezierTo(new Point(g1 - (w * 0.25), body.Bottom), new Point(gc + (w * 0.25), body.Bottom - depth), new Point(gc, body.Bottom - depth));
                path.CubicBezierTo(new Point(gc - (w * 0.25), body.Bottom - depth), new Point(g0 + (w * 0.25), body.Bottom), new Point(g0, body.Bottom));
            }

            path.LineTo(new Point(body.Left + radius, body.Bottom));
            path.ArcTo(new Point(body.Left, body.Bottom - radius), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
            path.LineTo(new Point(body.Left, body.Top + radius));
            path.ArcTo(new Point(body.Left + radius, body.Top), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
            path.EndFigure(isClosed: true);

            return geometry;
        }

        // ###########################################################################################
        // A ceramic disc, face on - FLAT, matte, never a shiny ball (owner report, 2026-11-09: "looks more like a round
        // bubble"): the disc a little wider than tall, its edge showing below and to the right as a coin's does, the face
        // lit evenly from the upper left with its rounded rim catching the light, the coating run down round each leg and
        // joined between them, the code printed on it.
        // ###########################################################################################
        private static void DrawDisc(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture)
        {
            Rect body = layout.Body;
            IBrush fill = PartPictureDrawing.Solid("#C9652B");
            IBrush edge = PartPictureDrawing.Solid("#8A3E14");
            Rect side = body.Translate(new Vector(2.5, 3.5));

            PartPictureDrawing.DrawLegs(context, layout);

            context.DrawRectangle(Brushes.Transparent, null, new RoundedRect(side, side.Width / 2, side.Height / 2), PartPictureDrawing.ShadowOf(body.Width * 0.12));

            if (layout.Tab is Rect skirt)
            {
                Geometry coating = PartPictureDrawing.DiscCoating(layout, skirt, body);
                context.DrawGeometry(fill, null, coating);
                context.DrawGeometry(PartPictureDrawing.Upright(), null, coating);
            }

            // The disc's thickness, seen past the face.
            context.DrawGeometry(edge, null, new EllipseGeometry(side));

            var face = new EllipseGeometry(body);
            context.DrawGeometry(fill, null, face);

            // Evenly lit, lighter to the upper left - no highlight spot.
            context.DrawGeometry(
                new LinearGradientBrush
                {
                    StartPoint = new RelativePoint(0.15, 0.05, RelativeUnit.Relative),
                    EndPoint = new RelativePoint(0.85, 0.95, RelativeUnit.Relative),
                    GradientStops =
                    {
                        new GradientStop(Color.FromArgb(0x2E, 0xFF, 0xFF, 0xFF), 0),
                        new GradientStop(Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF), 0.5),
                        new GradientStop(Color.FromArgb(0x26, 0x00, 0x00, 0x00), 1),
                    },
                },
                null,
                face);

            PartPictureDrawing.Print(context, picture.Marking, layout.Marking, "#2A1408", 15, bold: true);

            // The rounded rim: bright where it faces the light, dark where it turns away.
            context.DrawGeometry(
                null,
                new Pen(
                    new LinearGradientBrush
                    {
                        StartPoint = new RelativePoint(0.2, 0, RelativeUnit.Relative),
                        EndPoint = new RelativePoint(0.8, 1, RelativeUnit.Relative),
                        GradientStops =
                        {
                            new GradientStop(Color.FromArgb(0x80, 0xFF, 0xE6, 0xD2), 0),
                            new GradientStop(Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF), 0.45),
                            new GradientStop(Color.FromArgb(0x60, 0x40, 0x14, 0x00), 1),
                        },
                    },
                    2.5),
                new EllipseGeometry(body.Deflate(1.6)));

            context.DrawGeometry(null, new Pen(new SolidColorBrush(Color.FromArgb(0x90, 0x4A, 0x1C, 0x06)), 1), face);
        }

        // ###########################################################################################
        // The coating a dipped disc carries down its legs: round each leg's top, narrowing to the wire, and up between the
        // two legs to the disc - as the resin ran off it.
        // ###########################################################################################
        private static StreamGeometry DiscCoating(PartPictureGeometry.Layout layout, Rect skirt, Rect body)
        {
            var geometry = new StreamGeometry();
            Rect left = layout.Leads[0].Wire, right = layout.Leads[^1].Wire;
            double top = skirt.Top, bottom = skirt.Bottom, grip = left.Width / 2 + 1.4;
            double xl = left.Center.X, xr = right.Center.X, middle = (xl + xr) / 2;

            using StreamGeometryContext path = geometry.Open();
            path.BeginFigure(new Point(skirt.Left, top), isFilled: true);
            path.QuadraticBezierTo(new Point(xl - grip - 2, bottom - 7), new Point(xl - grip, bottom - 2));
            path.QuadraticBezierTo(new Point(xl, bottom + 1.5), new Point(xl + grip, bottom - 2));
            path.QuadraticBezierTo(new Point(xl + grip + 3, bottom - 9), new Point(middle, Math.Min(body.Bottom - 1, bottom - 8)));
            path.QuadraticBezierTo(new Point(xr - grip - 3, bottom - 9), new Point(xr - grip, bottom - 2));
            path.QuadraticBezierTo(new Point(xr, bottom + 1.5), new Point(xr + grip, bottom - 2));
            path.QuadraticBezierTo(new Point(xr + grip + 2, bottom - 7), new Point(skirt.Right, top));
            path.EndFigure(isClosed: true);

            return geometry;
        }

        // A dipped part standing on its legs - a radial ceramic capacitor, an EMI filter: a rounded drop of resin.
        private static void DrawDipped(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture, string color, string printColor)
        {
            Rect body = layout.Body;
            var drop = new RoundedRect(body, new CornerRadius(body.Height * 0.45, body.Height * 0.45, body.Height * 0.3, body.Height * 0.3));

            PartPictureDrawing.DrawLegs(context, layout);
            PartPictureDrawing.Shadow(context, drop, body.Height * 0.14);

            context.DrawRectangle(PartPictureDrawing.Solid(color), null, drop);
            context.DrawRectangle(PartPictureDrawing.Dome(), null, drop);

            PartPictureDrawing.Print(context, picture.Marking, layout.Marking, printColor, 12, bold: true);

            context.DrawRectangle(null, PartPictureDrawing.Edge(), drop);
        }

        // ###########################################################################################
        // A body lying across with nothing but its color and a band - an axial ceramic capacitor (its code printed), a
        // black DO-41 diode (its grey cathode band), a ferrite bead: a cylinder with rounded ends, lit from above.
        // ###########################################################################################
        private static void DrawAxialBody(
            DrawingContext context,
            PartPictureGeometry.Layout layout,
            PartPicture.Picture picture,
            string color,
            double radiusShare,
            string? markingColor,
            string? bandColor = null)
        {
            Rect body = layout.Body;
            var outline = new RoundedRect(body, body.Height * radiusShare);

            PartPictureDrawing.DrawLeads(context, layout);
            PartPictureDrawing.Shadow(context, outline, body.Height * 0.3);

            using (context.PushClip(outline))
            {
                context.DrawRectangle(PartPictureDrawing.Solid(color), null, body);

                if (bandColor is not null && layout.Band is Rect band)
                {
                    context.DrawRectangle(PartPictureDrawing.Solid(bandColor), null, band);
                }

                if (markingColor is not null)
                {
                    PartPictureDrawing.Print(context, picture.Marking.Replace('\n', ' '), layout.Marking, markingColor, 11, bold: true);
                }

                context.DrawRectangle(PartPictureDrawing.Lying(), null, body);
                PartPictureDrawing.RoundEnds(context, body);
            }

            context.DrawRectangle(null, PartPictureDrawing.Edge(), outline);
        }

        // A glass diode (DO-35): the lead seen through the orange glass, the black cathode band, a gleam along the top.
        private static void DrawGlassDiode(DrawingContext context, PartPictureGeometry.Layout layout)
        {
            Rect body = layout.Body;
            var outline = new RoundedRect(body, body.Height / 2);

            PartPictureDrawing.DrawLeads(context, layout);
            PartPictureDrawing.Shadow(context, outline, body.Height * 0.3);

            // The wire and its slugs through the glass.
            Rect wire = layout.Leads[0].Wire;
            PartPictureDrawing.DrawWire(context, new Rect(body.Left, wire.Top, body.Width, wire.Height), vertical: false);
            PartPictureDrawing.DrawWire(context, new Rect(body.Left + (body.Width * 0.18), body.Top + (body.Height * 0.3), body.Width * 0.64, body.Height * 0.4), vertical: false);

            using (context.PushClip(outline))
            {
                context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0xB8, 0xE2, 0x6A, 0x2A)), null, body);

                if (layout.Band is Rect band)
                {
                    context.DrawRectangle(PartPictureDrawing.Solid("#161616"), null, band);
                }

                context.DrawRectangle(PartPictureDrawing.Lying(), null, body);
                context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0xA0, 0xFF, 0xFF, 0xFF)), null, new RoundedRect(new Rect(body.Left + (body.Width * 0.14), body.Top + (body.Height * 0.16), body.Width * 0.72, body.Height * 0.12), 1));
                PartPictureDrawing.RoundEnds(context, body);
            }

            context.DrawRectangle(null, new Pen(new SolidColorBrush(Color.FromArgb(0x70, 0x5A, 0x20, 0x08)), 1), outline);
        }

        // A coil: copper wire wound between the two flanges of a dark ferrite core.
        private static void DrawCoil(DrawingContext context, PartPictureGeometry.Layout layout)
        {
            Rect body = layout.Body;
            double flange = body.Width * 0.1;
            var winding = new Rect(body.Left + flange - 1, body.Top + (body.Height * 0.1), body.Width - (2 * flange) + 2, body.Height * 0.8);
            var left = new RoundedRect(new Rect(body.Left, body.Top, flange, body.Height), 2);
            var right = new RoundedRect(new Rect(body.Right - flange, body.Top, flange, body.Height), 2);

            PartPictureDrawing.DrawLeads(context, layout);
            PartPictureDrawing.Shadow(context, new RoundedRect(body, 3), body.Height * 0.3);

            context.DrawRectangle(PartPictureDrawing.Solid("#B9733A"), null, winding);

            using (context.PushClip(winding))
            {
                IBrush turn = new SolidColorBrush(Color.FromArgb(0xA0, 0x5E, 0x2E, 0x0E));

                for (double x = winding.Left + 1.5; x < winding.Right; x += 3.2)
                {
                    context.DrawRectangle(turn, null, new Rect(x, winding.Top, 1.1, winding.Height));
                }
            }

            context.DrawRectangle(PartPictureDrawing.Lying(), null, winding);

            foreach (RoundedRect end in new[] { left, right })
            {
                context.DrawRectangle(PartPictureDrawing.Solid("#3E3E44"), null, end);
                context.DrawRectangle(PartPictureDrawing.Lying(), null, end);
                context.DrawRectangle(null, PartPictureDrawing.Edge(), end);
            }
        }

        // A TO-92 case from its flat, printed face: black, its top corners rounded, three thin legs.
        // ###########################################################################################
        // A TO-92 AS IT LOOKS (owner request, 2026-11-09: "it does not look like a transistor"): seen from its flat side and a
        // little from above - the top of the case, the D its round back makes, over the matte black face, the face's top
        // edge catching the light; the type printed in grey; three thin tinned legs out of its foot close together, bent out
        // and running down long (PartPictureGeometry.LegPath).
        // ###########################################################################################
        private static void DrawTO92(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture)
        {
            Rect body = layout.Body;
            Rect cap = layout.Tab ?? new Rect(body.Left, body.Top, body.Width, 0);
            var face = new RoundedRect(body, new CornerRadius(0, 0, 3, 3));

            PartPictureDrawing.DrawBentLegs(context, layout);
            PartPictureDrawing.Shadow(context, new RoundedRect(new Rect(body.Left, cap.Top, body.Width, body.Bottom - cap.Top), 4), body.Height * 0.14);

            // The top of the case: the back half of the ellipse shows, the face hides the front half.
            var top = new EllipseGeometry(cap);
            context.DrawGeometry(PartPictureDrawing.Solid("#2B2B2D"), null, top);
            context.DrawGeometry(
                new LinearGradientBrush
                {
                    StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                    EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative),
                    GradientStops =
                    {
                        new GradientStop(Color.FromArgb(0x40, 0, 0, 0), 0),
                        new GradientStop(Color.FromArgb(0x38, 0xFF, 0xFF, 0xFF), 0.3),
                        new GradientStop(Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF), 0.6),
                        new GradientStop(Color.FromArgb(0x50, 0, 0, 0), 1),
                    },
                },
                null,
                top);
            context.DrawGeometry(null, new Pen(new SolidColorBrush(Color.FromArgb(0x8C, 0x10, 0x10, 0x10)), 1), top);

            // The flat face, matte: barely lighter at its top, the print on it.
            context.DrawRectangle(PartPictureDrawing.Solid("#1A1A1B"), null, face);
            context.DrawRectangle(
                new LinearGradientBrush
                {
                    StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                    EndPoint = new RelativePoint(0, 1, RelativeUnit.Relative),
                    GradientStops =
                    {
                        new GradientStop(Color.FromArgb(0x1C, 0xFF, 0xFF, 0xFF), 0),
                        new GradientStop(Color.FromArgb(0x06, 0xFF, 0xFF, 0xFF), 0.5),
                        new GradientStop(Color.FromArgb(0x30, 0, 0, 0), 1),
                    },
                },
                null,
                face);
            PartPictureDrawing.Print(context, picture.Marking, layout.Marking, "#BDBDBD", 9.5);

            // The edge between the top and the face, catching the light.
            context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0x70, 0xFF, 0xFF, 0xFF)), null, new Rect(body.Left + 1, body.Top, body.Width - 2, 1));
            context.DrawRectangle(null, PartPictureDrawing.Edge(), face);
        }

        // ###########################################################################################
        // Legs drawn along their paths - a TO-92's bent out from the case - as round tinned wire: a soft shadow, the wire, a
        // bright line along its lit side.
        // ###########################################################################################
        private static void DrawBentLegs(DrawingContext context, PartPictureGeometry.Layout layout)
        {
            foreach (PartPictureGeometry.LeadPlace lead in layout.Leads)
            {
                IReadOnlyList<Point> path = PartPictureGeometry.LegPath(layout, lead);
                double thickness = lead.Wire.Width;

                PartPictureDrawing.Polyline(context, path, new Vector(1.2, 2), new Pen(new SolidColorBrush(Color.FromArgb(0x40, 0, 0, 0)), thickness + 1, lineCap: PenLineCap.Round, lineJoin: PenLineJoin.Round));
                PartPictureDrawing.Polyline(context, path, default, new Pen(PartPictureDrawing.Solid("#9A9A9A"), thickness, lineCap: PenLineCap.Round, lineJoin: PenLineJoin.Round));
                PartPictureDrawing.Polyline(context, path, new Vector(-thickness * 0.2, 0), new Pen(PartPictureDrawing.Solid("#EDEDED"), thickness * 0.35, lineCap: PenLineCap.Round, lineJoin: PenLineJoin.Round));
            }
        }

        private static void Polyline(DrawingContext context, IReadOnlyList<Point> points, Vector offset, Pen pen)
        {
            for (int i = 1; i < points.Count; i++)
            {
                context.DrawLine(pen, points[i - 1] + offset, points[i] + offset);
            }
        }

        // A TO-220 case from its printed face: the metal tab with its hole over the black body, three flat legs.
        private static void DrawTO220(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture)
        {
            Rect body = layout.Body;
            var block = new RoundedRect(body, 2);

            PartPictureDrawing.DrawLegs(context, layout);

            // Each leg wider where it leaves the body.
            foreach (Rect wire in layout.Leads.Select(lead => lead.Wire))
            {
                PartPictureDrawing.DrawWire(context, new Rect(wire.Left - 1.5, wire.Top, wire.Width + 3, 9), vertical: true);
            }

            if (layout.Tab is Rect tab && layout.Hole is Rect hole)
            {
                var plate = new CombinedGeometry(GeometryCombineMode.Exclude, new RectangleGeometry(tab), new EllipseGeometry(hole));
                PartPictureDrawing.Shadow(context, new RoundedRect(tab, 2), tab.Height * 0.1);

                context.DrawGeometry(
                    new LinearGradientBrush
                    {
                        StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                        EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
                        GradientStops =
                        {
                            new GradientStop(Color.Parse("#F2F3F5"), 0),
                            new GradientStop(Color.Parse("#B4B7BC"), 0.45),
                            new GradientStop(Color.Parse("#DADDE1"), 0.7),
                            new GradientStop(Color.Parse("#8F9399"), 1),
                        },
                    },
                    PartPictureDrawing.Edge(),
                    plate);

                context.DrawGeometry(null, new Pen(new SolidColorBrush(Color.FromArgb(0x90, 0x3A, 0x3A, 0x3A)), 1.5), new EllipseGeometry(hole.Deflate(0.5)));
            }

            PartPictureDrawing.Shadow(context, block, body.Height * 0.1);
            context.DrawRectangle(PartPictureDrawing.Solid("#1E1E1E"), null, block);
            context.DrawRectangle(PartPictureDrawing.Flat(), null, block);

            // The step where the body meets the tab.
            context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0x38, 0xFF, 0xFF, 0xFF)), null, new Rect(body.Left, body.Top, body.Width, 3));

            PartPictureDrawing.Print(context, picture.Marking, layout.Marking, "#C8C8C8", 11);
            context.DrawRectangle(null, PartPictureDrawing.Edge(), block);
        }

        // A crystal: the silver can on its base, the frequency stamped on it.
        private static void DrawCrystal(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture)
        {
            Rect body = layout.Body;
            var can = new RoundedRect(body, new CornerRadius(body.Width * 0.18, body.Width * 0.18, 0, 0));

            PartPictureDrawing.DrawLegs(context, layout);
            PartPictureDrawing.Shadow(context, can, body.Height * 0.14);

            if (layout.Tab is Rect foot)
            {
                context.DrawRectangle(PartPictureDrawing.Solid("#9EA2A8"), null, new RoundedRect(foot, 1.5));
                context.DrawRectangle(PartPictureDrawing.Upright(), null, new RoundedRect(foot, 1.5));
            }

            context.DrawRectangle(PartPictureDrawing.Solid("#D3D6DB"), null, can);
            context.DrawRectangle(PartPictureDrawing.Metal(), null, can);
            PartPictureDrawing.Print(context, picture.Marking, layout.Marking, "#3C3F44", 11, bold: true);
            context.DrawRectangle(null, PartPictureDrawing.Edge(), can);
        }

        // A resistor pack: a flat black block on its row of legs, the dot by pin 1, its value printed.
        private static void DrawResistorNetwork(DrawingContext context, PartPictureGeometry.Layout layout, PartPicture.Picture picture)
        {
            Rect body = layout.Body;
            var block = new RoundedRect(body, 4);

            PartPictureDrawing.DrawLegs(context, layout);
            PartPictureDrawing.Shadow(context, block, body.Height * 0.14);

            context.DrawRectangle(PartPictureDrawing.Solid("#2A2A2C"), null, block);
            context.DrawRectangle(PartPictureDrawing.Flat(), null, block);

            if (layout.Dot is Point dot)
            {
                context.DrawEllipse(PartPictureDrawing.Solid("#ECECEC"), null, dot, 3.5, 3.5);
            }

            PartPictureDrawing.Print(context, picture.Marking, layout.Marking, "#E6E6E6", 11.5);
            context.DrawRectangle(null, PartPictureDrawing.Edge(), block);
        }

        // A pin header: gold square posts through a black plastic base.
        private static void DrawPinHeader(DrawingContext context, PartPictureGeometry.Layout layout)
        {
            Rect body = layout.Body;

            foreach (PartPictureGeometry.LeadPlace lead in layout.Leads)
            {
                PartPictureDrawing.DrawGoldPin(context, lead.Wire);

                if (lead.Post is Rect post)
                {
                    PartPictureDrawing.DrawGoldPin(context, post);
                }
            }

            PartPictureDrawing.Shadow(context, new RoundedRect(body, 1.5), body.Height * 0.25);
            context.DrawRectangle(PartPictureDrawing.Solid("#1C1C1C"), null, new RoundedRect(body, 1.5));
            context.DrawRectangle(PartPictureDrawing.Flat(), null, new RoundedRect(body, 1.5));

            // The base is one block per pin, a groove between each two.
            for (int i = 1; i < layout.Leads.Count; i++)
            {
                double x = (layout.Leads[i - 1].Wire.Center.X + layout.Leads[i].Wire.Center.X) / 2;
                context.DrawRectangle(new SolidColorBrush(Color.FromArgb(0x90, 0, 0, 0)), null, new Rect(x - 0.75, body.Top, 1.5, body.Height));
            }

            context.DrawRectangle(null, PartPictureDrawing.Edge(), new RoundedRect(body, 1.5));
        }

        // ---- the pieces every part is made of -------------------------------------------------

        // The legs of a part standing, behind its body.
        private static void DrawLegs(DrawingContext context, PartPictureGeometry.Layout layout)
        {
            foreach (Rect wire in layout.Leads.Select(lead => lead.Wire))
            {
                PartPictureDrawing.Shadow(context, new RoundedRect(wire, wire.Width / 2), wire.Width * 0.8);
            }

            foreach (Rect wire in layout.Leads.Select(lead => lead.Wire))
            {
                PartPictureDrawing.DrawWire(context, wire, vertical: true);
            }
        }

        // The leads of a part lying across, behind its body.
        private static void DrawLeads(DrawingContext context, PartPictureGeometry.Layout layout)
        {
            foreach (Rect wire in layout.Leads.Select(lead => lead.Wire))
            {
                PartPictureDrawing.Shadow(context, new RoundedRect(wire, wire.Height / 2), wire.Height * 1.2);
                PartPictureDrawing.DrawWire(context, wire, vertical: false);
            }
        }

        // A lead: tinned copper wire, round - lit along its top (lying) or its left (standing), its end rounded.
        private static void DrawWire(DrawingContext context, Rect wire, bool vertical) =>
            context.DrawRectangle(
                new LinearGradientBrush
                {
                    StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                    EndPoint = vertical ? new RelativePoint(1, 0, RelativeUnit.Relative) : new RelativePoint(0, 1, RelativeUnit.Relative),
                    GradientStops =
                    {
                        new GradientStop(Color.Parse("#8C8C8C"), 0),
                        new GradientStop(Color.Parse("#F4F4F4"), 0.32),
                        new GradientStop(Color.Parse("#BDBDBD"), 0.6),
                        new GradientStop(Color.Parse("#6B6B6B"), 1),
                    },
                },
                null,
                new RoundedRect(wire, Math.Min(wire.Width, wire.Height) / 2));

        // A gold-plated square pin, lit from the left.
        private static void DrawGoldPin(DrawingContext context, Rect pin) =>
            context.DrawRectangle(
                new LinearGradientBrush
                {
                    StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                    EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative),
                    GradientStops =
                    {
                        new GradientStop(Color.Parse("#A9831E"), 0),
                        new GradientStop(Color.Parse("#F8DE8A"), 0.35),
                        new GradientStop(Color.Parse("#D3AA3C"), 0.65),
                        new GradientStop(Color.Parse("#8A6912"), 1),
                    },
                },
                null,
                new RoundedRect(pin, 0.8));

        // "-" signs down a radial electrolytic's stripe.
        private static void MinusMarks(DrawingContext context, Rect stripe, string color)
        {
            IBrush brush = PartPictureDrawing.Solid(color);
            const int count = 3;

            for (int i = 0; i < count; i++)
            {
                double length = stripe.Width * 0.5;
                double y = stripe.Top + (stripe.Height * (0.22 + (i * 0.25)));
                context.DrawRectangle(brush, null, new Rect(stripe.Center.X - (length / 2), y - 1.75, length, 3.5));
            }
        }

        // ###########################################################################################
        // What is printed on a body, a line per "\n", centred in `area`, made smaller until it fits. Drawn before the
        // body's shading, which goes over the print as over the paint.
        // ###########################################################################################
        private static void Print(DrawingContext context, string marking, Rect area, string color, double size, bool bold = false)
        {
            if (marking.Length == 0 || area.Width <= 0 || area.Height <= 0)
            {
                return;
            }

            IBrush brush = PartPictureDrawing.Solid(color);
            string[] lines = marking.Split('\n');
            FontWeight weight = bold ? FontWeight.Bold : FontWeight.SemiBold;

            List<FormattedText> texts = lines.Select(line => PartPictureDrawing.Text(line, brush, size, weight)).ToList();
            double widest = texts.Max(text => text.Width);
            double tall = texts.Sum(text => text.Height);
            double fit = Math.Min(1, Math.Min(area.Width / Math.Max(1, widest), area.Height / Math.Max(1, tall)));

            if (fit < 1)
            {
                texts = lines.Select(line => PartPictureDrawing.Text(line, brush, Math.Max(6, size * fit), weight)).ToList();
                tall = texts.Sum(text => text.Height);
            }

            double y = area.Center.Y - (tall / 2);

            foreach (FormattedText text in texts)
            {
                context.DrawText(text, new Point(area.Center.X - (text.Width / 2), y));
                y += text.Height;
            }
        }

        // A soft shadow under a body, `depth` saying how far it falls.
        private static void Shadow(DrawingContext context, RoundedRect shape, double depth) =>
            context.DrawRectangle(Brushes.Transparent, null, shape, PartPictureDrawing.ShadowOf(depth));

        private static BoxShadows ShadowOf(double depth) => new(new BoxShadow
        {
            OffsetX = depth * 0.25,
            OffsetY = depth * 0.5,
            Blur = depth,
            Color = Color.FromArgb(0x59, 0x00, 0x00, 0x00),
        });

        // A body lying across, round: lit from above, a highlight a quarter down, darker under it (the resistor's).
        private static LinearGradientBrush Lying() => PartPictureDrawing.Round(new RelativePoint(0, 1, RelativeUnit.Relative));

        // A body standing, round: lit from the left.
        private static LinearGradientBrush Upright() => PartPictureDrawing.Round(new RelativePoint(1, 0, RelativeUnit.Relative));

        private static LinearGradientBrush Round(RelativePoint end) => new()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = end,
            GradientStops =
            {
                new GradientStop(Color.FromArgb(0x40, 0x00, 0x00, 0x00), 0),
                new GradientStop(Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF), 0.1),
                new GradientStop(Color.FromArgb(0x80, 0xFF, 0xFF, 0xFF), 0.25),
                new GradientStop(Color.FromArgb(0x1F, 0xFF, 0xFF, 0xFF), 0.38),
                new GradientStop(Color.FromArgb(0x00, 0x00, 0x00, 0x00), 0.55),
                new GradientStop(Color.FromArgb(0x2E, 0x00, 0x00, 0x00), 0.8),
                new GradientStop(Color.FromArgb(0x73, 0x00, 0x00, 0x00), 1),
            },
        };

        // A flat face, lit from above: a sheen at its top, a little darker at its foot.
        private static LinearGradientBrush Flat() => new()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(0, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.FromArgb(0x38, 0xFF, 0xFF, 0xFF), 0),
                new GradientStop(Color.FromArgb(0x10, 0xFF, 0xFF, 0xFF), 0.3),
                new GradientStop(Color.FromArgb(0x00, 0x00, 0x00, 0x00), 0.6),
                new GradientStop(Color.FromArgb(0x30, 0x00, 0x00, 0x00), 1),
            },
        };

        // Polished metal, standing: bright bands down it, as a can catches the light.
        private static LinearGradientBrush Metal() => new()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.FromArgb(0x50, 0x00, 0x00, 0x00), 0),
                new GradientStop(Color.FromArgb(0x90, 0xFF, 0xFF, 0xFF), 0.22),
                new GradientStop(Color.FromArgb(0x10, 0xFF, 0xFF, 0xFF), 0.42),
                new GradientStop(Color.FromArgb(0x26, 0x00, 0x00, 0x00), 0.62),
                new GradientStop(Color.FromArgb(0x60, 0xFF, 0xFF, 0xFF), 0.8),
                new GradientStop(Color.FromArgb(0x60, 0x00, 0x00, 0x00), 1),
            },
        };

        // A rounded face - a disc, a drop of resin: lit at its upper left, darker towards its rim.
        private static RadialGradientBrush Dome() => new()
        {
            Center = new RelativePoint(0.38, 0.32, RelativeUnit.Relative),
            GradientOrigin = new RelativePoint(0.34, 0.26, RelativeUnit.Relative),
            RadiusX = new RelativeScalar(0.75, RelativeUnit.Relative),
            RadiusY = new RelativeScalar(0.75, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.FromArgb(0x70, 0xFF, 0xFF, 0xFF), 0),
                new GradientStop(Color.FromArgb(0x10, 0xFF, 0xFF, 0xFF), 0.35),
                new GradientStop(Color.FromArgb(0x00, 0x00, 0x00, 0x00), 0.55),
                new GradientStop(Color.FromArgb(0x50, 0x00, 0x00, 0x00), 0.9),
                new GradientStop(Color.FromArgb(0x70, 0x00, 0x00, 0x00), 1),
            },
        };

        // The rounded ends of a body lying across, a little darker.
        private static void RoundEnds(DrawingContext context, Rect body)
        {
            context.DrawRectangle(PartPictureDrawing.DarkerAtEdge(fromLeft: true), null, body);
            context.DrawRectangle(PartPictureDrawing.DarkerAtEdge(fromLeft: false), null, body);
        }

        private static LinearGradientBrush DarkerAtEdge(bool fromLeft) => new()
        {
            StartPoint = new RelativePoint(fromLeft ? 0 : 1, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(fromLeft ? 1 : 0, 0, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.FromArgb(0x38, 0, 0, 0), 0),
                new GradientStop(Color.FromArgb(0, 0, 0, 0), 0.12),
            },
        };

        private static Pen Edge() => new(new SolidColorBrush(Color.FromArgb(0x8C, 0x20, 0x20, 0x20)), 1);

        private static IBrush Solid(string color) => new SolidColorBrush(Color.Parse(color));

        private IBrush Brush(string key, IBrush fallback) =>
            this.TryFindResource(key, this.ActualThemeVariant, out object? value) && value is IBrush brush ? brush : fallback;

        private static FormattedText Text(string text, IBrush brush, double size, FontWeight weight, double maxWidth = double.PositiveInfinity) =>
            new(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(FontFamily.Default, FontStyle.Normal, weight), size, brush)
            {
                MaxTextWidth = maxWidth,
                MaxLineCount = 1,
                Trimming = TextTrimming.CharacterEllipsis,
            };
    }
}
