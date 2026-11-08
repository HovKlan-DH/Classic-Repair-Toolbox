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
    // A RESISTOR, DRAWN WITH ITS COLOR BANDS, under the pinout in the component popup (owner request, 2026-11-07, with
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

        // The bands drawn - none hides nothing by itself; the popup hides the section.
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

            var leadFill = new SolidColorBrush(Color.Parse("#CFCFCF"));
            var leadEdge = new Pen(new SolidColorBrush(Color.Parse("#A6A6A6")), 1);
            context.DrawRectangle(leadFill, leadEdge, layout.LeftLead);
            context.DrawRectangle(leadFill, leadEdge, layout.RightLead);

            Geometry body = BodyOutline(layout);
            var bodyFill = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                EndPoint = new RelativePoint(0, 1, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(Color.Parse("#F0EADF"), 0),
                    new GradientStop(Color.Parse("#E3DACB"), 0.6),
                    new GradientStop(Color.Parse("#D3C8B6"), 1),
                },
            };

            context.DrawGeometry(bodyFill, null, body);

            using (context.PushGeometryClip(body))
            {
                for (int i = 0; i < this.thisBands.Count; i++)
                {
                    context.DrawRectangle(BrushOf(this.thisBands[i].Color), null, layout.Bands[i]);
                }
            }

            context.DrawGeometry(null, new Pen(new SolidColorBrush(Color.Parse("#9C968C")), 1.2), body);

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
