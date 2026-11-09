using Avalonia;
using Avalonia.Controls;
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
    // A RESISTOR, DRAWN WITH ITS COLOR BANDS, in the pinout's place - the component popup's and the "Pinout" page's, as
    // ResistorPicture shows it with its value and colors (owner requests, 2026-11-07, and 2026-11-09: "the image/drawing of
    // the resistor should be in the pinout instead (like the other components you are drawing)"; with
    // the picture the owner made by hand per value - "560 Ohms 5%": a beige body with wider end caps, the lead stubs,
    // the bands, and a faint stem from each band to a pill in its color naming what it stands for). The bands are
    // ResistorColorCode's, every position ResistorBandsGeometry's; this only paints. Its own colors, a picture of the
    // part, on either theme.
    // ###########################################################################################
    internal sealed class ResistorBandsDrawing : Control
    {
        // The widest it is drawn - the popup's other drawings' size.
        public const double MaxDrawWidth = 320;

        private const double PillFontSize = 10.5;

        private IReadOnlyList<ResistorColorCode.ColorBand> thisBands = [];

        // The bands drawn - none hides nothing by itself; ResistorPicture hides itself.
        public IReadOnlyList<ResistorColorCode.ColorBand> Bands
        {
            get => this.thisBands;
            set
            {
                this.thisBands = value ?? [];
                this.InvalidateMeasure();
                this.InvalidateVisual();
            }
        }

        private bool WithTolerance => this.thisBands.Count > 0 && this.thisBands[^1].Kind == ResistorColorCode.BandKind.Tolerance;

        protected override Size MeasureOverride(Size availableSize)
        {
            double width = double.IsPositiveInfinity(availableSize.Width) ? MaxDrawWidth : Math.Min(MaxDrawWidth, availableSize.Width);
            return new Size(width, ResistorBandsGeometry.For(width, this.thisBands.Count, this.WithTolerance).Height);
        }

        public override void Render(DrawingContext context)
        {
            base.Render(context);

            double width = this.Bounds.Width;

            if (width <= 1 || this.thisBands.Count == 0)
            {
                return;
            }

            ResistorBandsGeometry.Layout layout = ResistorBandsGeometry.For(width, this.thisBands.Count, this.WithTolerance);
            Geometry body = BodyOutline(layout);

            DrawShadow(context, layout);
            DrawLead(context, layout.LeftLead);
            DrawLead(context, layout.RightLead);

            // Metal film light blue, carbon film beige (ResistorColorCode.IsMetalFilm).
            context.DrawGeometry(new SolidColorBrush(Color.Parse(ResistorColorCode.IsMetalFilm(this.thisBands) ? "#93C3EA" : "#E2CFAA")), null, body);

            using (context.PushGeometryClip(body))
            {
                for (int i = 0; i < this.thisBands.Count; i++)
                {
                    context.DrawRectangle(BrushOf(this.thisBands[i].Color), null, layout.Bands[i]);
                }

                DrawRoundness(context, layout);
            }

            context.DrawGeometry(null, new Pen(new SolidColorBrush(Color.FromArgb(0x8C, 0x5E, 0x50, 0x3A)), 1), body);

            // The pills, measured first, placed apart.
            var texts = this.thisBands
                .Select(band => new FormattedText(
                    band.Label,
                    CultureInfo.InvariantCulture,
                    FlowDirection.LeftToRight,
                    new Typeface(FontFamily.Default, FontStyle.Normal, FontWeight.SemiBold),
                    PillFontSize,
                    TextBrushOf(band.Color)))
                .ToList();

            List<double> widths = texts.Select(text => Math.Max(22, Math.Ceiling(text.Width) + 14)).ToList();
            IReadOnlyList<double> lefts = ResistorBandsGeometry.PillLefts(layout.Bands.Select(band => band.Center.X).ToList(), widths, width);

            for (int i = 0; i < this.thisBands.Count; i++)
            {
                Rect band = layout.Bands[i];
                Color fill = ColorOf(this.thisBands[i].Color);

                // The stem: the band's color, faint, from the body's foot to the pill.
                double stemTop = ResistorBandsGeometry.StemTop(layout, band);
                var stem = new Rect(band.X + band.Width * 0.1, stemTop, band.Width * 0.8, Math.Max(0, layout.PillTop - stemTop));
                context.DrawRectangle(new SolidColorBrush(fill, 0.22), null, stem);

                var pill = new Rect(lefts[i], layout.PillTop, widths[i], ResistorBandsGeometry.PillHeight);
                context.DrawRectangle(new SolidColorBrush(fill), new Pen(new SolidColorBrush(Color.FromArgb(0x55, 0, 0, 0)), 1), pill, 3, 3);

                FormattedText text = texts[i];
                context.DrawText(text, new Point(pill.X + (pill.Width - text.Width) / 2, pill.Y + (pill.Height - text.Height) / 2));
            }
        }

        // ###########################################################################################
        // *** A PART LYING ON THE PAGE, NOT A FLAT SHAPE (owner request, 2026-11-08: "make it more 'real', so it looks a
        // little more like a real resistor? Maybe some depth, and shadow"). *** A soft shadow under the body and the leads;
        // the leads round wire; the body - bands and all, as they are painted on it - shaded as the cylinder it is: lit
        // from above, a highlight a quarter down, darker towards its underside and at its rounded ends. The narrower middle
        // is shaded over its own height, so the caps stand proud of it.
        // ###########################################################################################
        private static void DrawShadow(DrawingContext context, ResistorBandsGeometry.Layout layout)
        {
            Rect body = layout.Body;
            double height = body.Height;
            var shadow = new BoxShadows(new BoxShadow
            {
                OffsetX = 0,
                OffsetY = height * 0.16,
                Blur = height * 0.32,
                Color = Color.FromArgb(0x59, 0x00, 0x00, 0x00),
            });

            var leads = new Rect(layout.LeftLead.Left, layout.LeftLead.Top, layout.RightLead.Right - layout.LeftLead.Left, layout.LeftLead.Height);
            var leadShadow = new BoxShadows(new BoxShadow
            {
                OffsetX = 0,
                OffsetY = height * 0.12,
                Blur = height * 0.12,
                Color = Color.FromArgb(0x40, 0x00, 0x00, 0x00),
            });

            context.DrawRectangle(Brushes.Transparent, null, new RoundedRect(leads, leads.Height / 2), leadShadow);
            context.DrawRectangle(Brushes.Transparent, null, new RoundedRect(body.Deflate(new Thickness(height * 0.1, 0)), height * 0.3), shadow);
        }

        // A lead: tinned copper wire, round - lit along its top, its end rounded.
        private static void DrawLead(DrawingContext context, Rect lead)
        {
            var wire = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                EndPoint = new RelativePoint(0, 1, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(Color.Parse("#8C8C8C"), 0),
                    new GradientStop(Color.Parse("#F4F4F4"), 0.32),
                    new GradientStop(Color.Parse("#BDBDBD"), 0.6),
                    new GradientStop(Color.Parse("#6B6B6B"), 1),
                },
            };

            context.DrawRectangle(wire, null, new RoundedRect(lead, lead.Height / 2));
        }

        private static void DrawRoundness(DrawingContext context, ResistorBandsGeometry.Layout layout)
        {
            Rect body = layout.Body;
            double cap = layout.CapWidth;
            double inset = layout.MiddleInset;

            var middle = new Rect(body.Left + cap, body.Top + inset, Math.Max(0, body.Width - 2 * cap), Math.Max(0, body.Height - 2 * inset));
            var leftCap = new Rect(body.Left, body.Top, cap, body.Height);
            var rightCap = new Rect(body.Right - cap, body.Top, cap, body.Height);

            // Round across: each part over its own height.
            foreach (Rect part in new[] { middle, leftCap, rightCap })
            {
                context.DrawRectangle(Cylinder(), null, part);
            }

            // Round at its ends, and the middle a step down from each cap. The ends only a little darker (owner request,
            // 2026-11-08: darker, they "could maybe be confused to be a color ring").
            context.DrawRectangle(DarkerAtEdge(0x1F, fromLeft: true, reach: 0.22), null, leftCap);
            context.DrawRectangle(DarkerAtEdge(0x1F, fromLeft: false, reach: 0.22), null, rightCap);
            context.DrawRectangle(DarkerAtEdge(0x38, fromLeft: true, reach: 0.035), null, middle);
            context.DrawRectangle(DarkerAtEdge(0x38, fromLeft: false, reach: 0.035), null, middle);
        }

        // Light from above, over whatever is under it - the body's color or a band's.
        private static LinearGradientBrush Cylinder() => new()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(0, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.FromArgb(0x40, 0x00, 0x00, 0x00), 0),
                new GradientStop(Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF), 0.1),
                new GradientStop(Color.FromArgb(0x9E, 0xFF, 0xFF, 0xFF), 0.25),
                new GradientStop(Color.FromArgb(0x1F, 0xFF, 0xFF, 0xFF), 0.38),
                new GradientStop(Color.FromArgb(0x00, 0x00, 0x00, 0x00), 0.55),
                new GradientStop(Color.FromArgb(0x2E, 0x00, 0x00, 0x00), 0.8),
                new GradientStop(Color.FromArgb(0x73, 0x00, 0x00, 0x00), 1),
            },
        };

        // Black of `alpha` at one edge, gone `reach` of the width in from it.
        private static LinearGradientBrush DarkerAtEdge(byte alpha, bool fromLeft, double reach) => new()
        {
            StartPoint = new RelativePoint(fromLeft ? 0 : 1, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(fromLeft ? 1 : 0, 0, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.FromArgb(alpha, 0, 0, 0), 0),
                new GradientStop(Color.FromArgb(0, 0, 0, 0), reach),
            },
        };

        // ###########################################################################################
        // The body's outline: the two end caps full height with rounded corners, the middle a little narrower, the
        // steps between them eased - as a carbon-film resistor looks.
        // ###########################################################################################
        private static Geometry BodyOutline(ResistorBandsGeometry.Layout layout)
        {
            Rect body = layout.Body;
            double cap = layout.CapWidth;
            double inset = layout.MiddleInset;
            double taper = cap * 0.22;
            double radius = Math.Min(body.Height * 0.18, cap * 0.4);

            var geometry = new StreamGeometry();

            using (StreamGeometryContext path = geometry.Open())
            {
                double left = body.Left, right = body.Right, top = body.Top, bottom = body.Bottom;

                path.BeginFigure(new Point(left + radius, top), isFilled: true);
                path.LineTo(new Point(left + cap - taper, top));
                path.QuadraticBezierTo(new Point(left + cap, top), new Point(left + cap, top + inset));
                path.LineTo(new Point(right - cap, top + inset));
                path.QuadraticBezierTo(new Point(right - cap, top), new Point(right - cap + taper, top));
                path.LineTo(new Point(right - radius, top));
                path.ArcTo(new Point(right, top + radius), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
                path.LineTo(new Point(right, bottom - radius));
                path.ArcTo(new Point(right - radius, bottom), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
                path.LineTo(new Point(right - cap + taper, bottom));
                path.QuadraticBezierTo(new Point(right - cap, bottom), new Point(right - cap, bottom - inset));
                path.LineTo(new Point(left + cap, bottom - inset));
                path.QuadraticBezierTo(new Point(left + cap, bottom), new Point(left + cap - taper, bottom));
                path.LineTo(new Point(left + radius, bottom));
                path.ArcTo(new Point(left, bottom - radius), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
                path.LineTo(new Point(left, top + radius));
                path.ArcTo(new Point(left + radius, top), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
                path.EndFigure(isClosed: true);
            }

            return geometry;
        }

        private static Color ColorOf(ResistorColorCode.Band band) => Color.Parse(ResistorColorCode.ColorsOf(band).Fill);

        private static IBrush BrushOf(ResistorColorCode.Band band) => new SolidColorBrush(ColorOf(band));

        private static IBrush TextBrushOf(ResistorColorCode.Band band) => new SolidColorBrush(Color.Parse(ResistorColorCode.ColorsOf(band).Text));
    }
}
