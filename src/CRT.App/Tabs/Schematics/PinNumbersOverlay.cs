using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Tabs.TabSchematics
{
    // A pin's number to draw on its pad: where the pad is in the image's own coordinates, how large it is there, its number,
    // whether it is pin 1, and the color of its direction - null to draw it plain.
    public sealed record PinNumberBadge(Point Center, double PadSize, string Number, bool IsFirst, PinDirectionKind? Direction);

    // Pin 1's orange pad, as the KiCad overlay draws it, to frame: its middle and size in the image's own coordinates, its turn
    // (clockwise degrees, as the overlay's) and whether it is square-cornered or round.
    public sealed record FirstPinMark(Point Center, Size Size, double RotationDegrees, bool IsRectangle);

    // ###########################################################################################
    // EVERY PIN OF A COMPONENT NUMBERED ON ITS PAD WHILE SHIFT IS HELD (owner request, 2026-11-17: "a new checkbox with the
    // text "Number all pins holding SHIFT", and if it is checked, then it will insert a pin number badge for all component
    // pins, and ideally that badge will be placed directly on top of the pin in the UI ... do visualize pin 1 specifically,
    // so it stands out from the others"; TabSchematics.PinsOnImage.cs decides which). Drawn in the panel's own coordinates -
    // never under the view's matrix - so a number stays as large as it reads at any zoom: a badge as tall as its pad is on
    // screen, within PinBadgeGeometry's least and most.
    //
    // PLAIN, each badge is the grey of a pin with no direction (owner report, 2026-11-17: "the pin color for this is wrong, and
    // you should revert to the previous color, which I think was grey'ish" - the hover badges' khaki until then) and pin 1's
    // the first pin's orange, its number bold. IN ITS DIRECTION'S COLORS ("Show color for pin direction" - "identical coloring
    // as from its pinout"; PinDirectionColors), pin 1 keeps its direction's colors and is told another way (owner request, the
    // same day: "pin 1 should not be specifically highlighted ... unless it has "another way" of visualizing if this is pin
    // 1"): a ring in the first pin's orange round it, 3 wide with a black line each side of it (owner report, the same day:
    // "the pin 1 marking is quite hard to see ... Maybe orange square should have a black border around it on both sides, and
    // the orange width should be +1px?"), its number bold.
    //
    // AND PIN 1'S ORANGE PAD, NOT NUMBERED, FRAMED IN BLACK (owner request, 2026-11-17: "can we then also have the "normal"
    // first pin marking having this 1px 2xblack border so it is more visible?", then "I do not want the inner black border -
    // just the one outer black border is fine"): a black line just outside the orange the KiCad overlay draws - here, in the
    // panel's own pixels, as the overlay is drawn in the image's and its lines would thin out to nothing at the first view
    // and thicken as it is zoomed.
    // ###########################################################################################
    public sealed class PinNumbersOverlay : Control
    {
        private IReadOnlyList<PinNumberBadge> thisBadges = [];

        private IReadOnlyList<FirstPinMark> thisFirstPins = [];

        public Matrix ViewMatrix { get; set; } = Matrix.Identity;

        public IReadOnlyList<PinNumberBadge> Badges => this.thisBadges;

        public IReadOnlyList<FirstPinMark> FirstPins => this.thisFirstPins;

        public void Show(IReadOnlyList<PinNumberBadge> badges, IReadOnlyList<FirstPinMark>? firstPins = null)
        {
            this.thisBadges = badges;
            this.thisFirstPins = firstPins ?? [];
            this.InvalidateVisual();
        }

        public void Clear()
        {
            if (this.thisBadges.Count > 0 || this.thisFirstPins.Count > 0)
            {
                this.thisBadges = [];
                this.thisFirstPins = [];
                this.InvalidateVisual();
            }
        }

        public override void Render(DrawingContext context)
        {
            this.ForgetWhatAnotherThemePainted();

            double scale = this.ViewMatrix.M11 <= 0 ? 1 : this.ViewMatrix.M11;
            var bounds = new Rect(this.Bounds.Size);
            IBrush firstPin = this.Brush("Schematics_FirstPin", Brushes.Orange);

            foreach (FirstPinMark mark in this.thisFirstPins)
            {
                PinNumbersOverlay.FrameFirstPin(context, mark, this.ViewMatrix, scale, bounds);
            }

            foreach (PinNumberBadge badge in this.thisBadges)
            {
                Point center = badge.Center.Transform(this.ViewMatrix);
                double height = PinBadgeGeometry.HeightFor(badge.PadSize * scale);
                double textSize = PinBadgeGeometry.TextSizeFor(height);

                // Off screen judged before any text is made, by the widest its number can be - a character never wider than
                // its size: zoomed in, most of a large selection's badges lie outside the panel at every frame.
                if (!PinBadgeGeometry.AreaOf(center, height, badge.Number.Length * textSize).Inflate(4).Intersects(bounds))
                {
                    continue;
                }

                FormattedText number = this.TextOf(badge, textSize);

                Rect area = PinBadgeGeometry.AreaOf(center, height, number.Width);

                if (!area.Inflate(4).Intersects(bounds))
                {
                    continue;
                }

                if (badge.Direction is { } direction)
                {
                    context.DrawRectangle(
                        this.Brush(PinDirectionColors.FillKeyOf(direction), Brushes.WhiteSmoke),
                        this.PenOf(PinDirectionColors.InkKeyOf(direction), Brushes.Gray, 1),
                        new RoundedRect(area, 3));

                    // Pin 1: a ring 3 wide in the first pin's orange, a black line each side of it, so it reads on any board.
                    if (badge.IsFirst)
                    {
                        context.DrawRectangle(null, OutlinePen, new RoundedRect(area.Inflate(1), 4));
                        context.DrawRectangle(null, this.PenOf("Schematics_FirstPin", Brushes.Orange, 3), new RoundedRect(area.Inflate(3), 6));
                        context.DrawRectangle(null, OutlinePen, new RoundedRect(area.Inflate(5), 8));
                    }
                }
                else
                {
                    // The grey of a pin with no direction - pin 1 in the first pin's orange - edged in black.
                    context.DrawRectangle(
                        badge.IsFirst ? firstPin : this.Brush(PinDirectionColors.FillKeyOf(PinDirectionKind.Other), Brushes.WhiteSmoke),
                        OutlinePen,
                        new RoundedRect(area, 3));
                }

                context.DrawText(number, new Point(area.Center.X - (number.Width / 2), area.Center.Y - (number.Height / 2)));
            }
        }

        private static readonly Pen OutlinePen = new(Brushes.Black, 1);

        // ###########################################################################################
        // A badge's number laid out once and KEPT while its number, weight, size, color and the theme are the same (code
        // review, 2026-11-18: a thousand laid out again at every frame of a pan or zoom with SHIFT held). Let go of whole when
        // it grows past KeptAtMost - each zoom step's size makes new ones.
        // ###########################################################################################
        private const int KeptAtMost = 4096;

        private readonly Dictionary<(string Number, bool IsFirst, double Size, IBrush Brush), FormattedText> thisTexts = [];

        // The theme's brushes and the pens made of them, found once per theme (code review, 2026-11-18: looked up in the
        // resources and made again for every badge at every frame).
        private readonly Dictionary<string, IBrush> thisBrushes = new(StringComparer.Ordinal);

        private readonly Dictionary<(string Key, double Thickness), Pen> thisPens = [];

        private Avalonia.Styling.ThemeVariant? thisPaintedTheme;

        // What another theme painted - its texts' colors, its brushes and pens - let go of as the theme changes.
        private void ForgetWhatAnotherThemePainted()
        {
            if (!Equals(this.thisPaintedTheme, this.ActualThemeVariant))
            {
                this.thisTexts.Clear();
                this.thisBrushes.Clear();
                this.thisPens.Clear();
                this.thisPaintedTheme = this.ActualThemeVariant;
            }
        }

        private Pen PenOf(string key, IBrush fallback, double thickness)
        {
            if (!this.thisPens.TryGetValue((key, thickness), out Pen? pen))
            {
                pen = new Pen(this.Brush(key, fallback), thickness);
                this.thisPens[(key, thickness)] = pen;
            }

            return pen;
        }

        private FormattedText TextOf(PinNumberBadge badge, double size)
        {
            if (this.thisTexts.Count > KeptAtMost)
            {
                this.thisTexts.Clear();
            }

            IBrush brush = this.NumberBrushOf(badge);
            var key = (badge.Number, badge.IsFirst, size, brush);

            if (!this.thisTexts.TryGetValue(key, out FormattedText? text))
            {
                text = new FormattedText(
                    badge.Number,
                    CultureInfo.InvariantCulture,
                    FlowDirection.LeftToRight,
                    new Typeface(FontFamily.Default, weight: badge.IsFirst ? FontWeight.Bold : FontWeight.Normal),
                    size,
                    brush);

                this.thisTexts[key] = text;
            }

            return text;
        }

        // ###########################################################################################
        // Pin 1's orange pad framed: the KiCad overlay draws it at its size with a 1.2 outline of its own, both in the image's
        // coordinates - so its orange ends 0.6 of them, at this zoom, outside its size. A black line just outside that, turned
        // with the pad.
        // ###########################################################################################
        private static void FrameFirstPin(DrawingContext context, FirstPinMark mark, Matrix view, double scale, Rect bounds)
        {
            Point centre = mark.Center.Transform(view);
            Rect outer = PinBadgeGeometry.FirstPinFrame(centre, mark.Size, scale);

            if (!outer.Inflate(4).Intersects(bounds))
            {
                return;
            }

            using DrawingContext.PushedState turned = KiCadPadGeometry.IsAxisAligned(mark.RotationDegrees)
                ? context.PushTransform(Matrix.Identity)
                : context.PushTransform(
                    Matrix.CreateTranslation(-centre.X, -centre.Y) *
                    Matrix.CreateRotation(mark.RotationDegrees * System.Math.PI / 180.0) *
                    Matrix.CreateTranslation(centre.X, centre.Y));

            if (mark.IsRectangle)
            {
                context.DrawRectangle(null, OutlinePen, outer);
            }
            else
            {
                context.DrawEllipse(null, OutlinePen, outer.Center, outer.Width / 2, outer.Height / 2);
            }
        }

        private IBrush NumberBrushOf(PinNumberBadge badge) =>
            badge.Direction is { } direction
                ? this.Brush(PinDirectionColors.NumberKeyOf(direction), Brushes.Black)
                : badge.IsFirst ? Brushes.Black : this.Brush(PinDirectionColors.NumberKeyOf(PinDirectionKind.Other), Brushes.Black);

        private IBrush Brush(string key, IBrush fallback)
        {
            if (!this.thisBrushes.TryGetValue(key, out IBrush? brush))
            {
                brush = this.TryFindResource(key, this.ActualThemeVariant, out object? value) && value is IBrush found ? found : fallback;
                this.thisBrushes[key] = brush;
            }

            return brush;
        }
    }
}
